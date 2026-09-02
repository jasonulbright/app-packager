<#
.SYNOPSIS
    Core module for Installer Analysis Tool.

.DESCRIPTION
    Provides functions for analyzing installer files (EXE, MSI):
      - Structured logging via the vendored SuiteCommon module (Lib\SuiteCommon)
      - File identification (version info, PE architecture, digital signature)
      - Installer type detection (MSI, NSIS, Inno Setup, InstallShield, WiX Burn, etc.)
      - MSI property extraction (via PSGallery MSI module or COM fallback)
      - Silent install switch lookup
      - Payload extraction via 7-Zip
      - Binary string analysis
      - Export to CSV, HTML, clipboard

.EXAMPLE
    Import-Module "$PSScriptRoot\Module\InstallerAnalysisCommon.psd1" -Force
    $info = Get-InstallerFileInfo -Path "C:\temp\setup.exe"
    $type = Get-InstallerType -Path "C:\temp\setup.exe"
#>

# ---------------------------------------------------------------------------
# Shared core (vendored SuiteCommon)
# ---------------------------------------------------------------------------
# Logging (Initialize-Logging, Write-Log) and settings persistence come from
# the vendored copy at Lib\SuiteCommon\. -Global makes the functions
# resolvable from the shell script and from this module alike; the guard
# keeps a -Force reimport of this module from resetting SuiteCommon state
# mid-session.
if (-not (Get-Module SuiteCommon)) {
    Import-Module (Join-Path $PSScriptRoot '..\Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking
}

# ---------------------------------------------------------------------------
# Module-scoped state
# ---------------------------------------------------------------------------

$script:MsiModuleAvailable = $null

# ---------------------------------------------------------------------------
# File Identification
# ---------------------------------------------------------------------------

function Get-PeArchitecture {
    <#
    .SYNOPSIS
        Reads PE header to determine architecture (x86/x64/ARM64).
    #>
    param([Parameter(Mandatory)][string]$Path)

    $stream = $null
    $reader = $null
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        $reader = New-Object System.IO.BinaryReader($stream)

        $mz = $reader.ReadUInt16()
        if ($mz -ne 0x5A4D) { return 'Not a PE' }

        $stream.Seek(0x3C, [System.IO.SeekOrigin]::Begin) | Out-Null
        $peOffset = $reader.ReadInt32()

        $stream.Seek($peOffset, [System.IO.SeekOrigin]::Begin) | Out-Null
        $peSignature = $reader.ReadUInt32()
        if ($peSignature -ne 0x00004550) { return 'Invalid PE' }

        $machineType = $reader.ReadUInt16()
        switch ($machineType) {
            0x014C  { 'x86' }
            0x8664  { 'x64' }
            0xAA64  { 'ARM64' }
            0x01C0  { 'ARM' }
            default { "Unknown (0x$($machineType.ToString('X4')))" }
        }
    }
    catch {
        return 'Error'
    }
    finally {
        if ($reader) { try { $reader.Close() } catch { $null = $_ } }
        if ($stream) { try { $stream.Close() } catch { $null = $_ } }
    }
}

function Get-InstallerFileInfo {
    <#
    .SYNOPSIS
        Extracts file metadata: version info, PE architecture, digital signature, size, hash.
    #>
    param([Parameter(Mandatory)][string]$Path)

    Write-Log "Analyzing file: $Path"

    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Path)
    $arch = if ($item.Extension -eq '.msi') { 'N/A (see MSI Summary)' } else { Get-PeArchitecture -Path $Path }
    $sig = Get-AuthenticodeSignature -FilePath $Path -ErrorAction SilentlyContinue
    $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction SilentlyContinue).Hash

    return [PSCustomObject]@{
        FileName         = $item.Name
        FullPath         = $item.FullName
        Extension        = $item.Extension
        FileSize         = $item.Length
        FileSizeFormatted = if ($item.Length -ge 1MB) { "{0:N1} MB" -f ($item.Length / 1MB) } else { "{0:N0} KB" -f ($item.Length / 1KB) }
        SHA256           = $hash
        FileVersion      = $versionInfo.FileVersion
        ProductVersion   = $versionInfo.ProductVersion
        ProductName      = $versionInfo.ProductName
        CompanyName      = $versionInfo.CompanyName
        FileDescription  = $versionInfo.FileDescription
        OriginalFilename = $versionInfo.OriginalFilename
        LegalCopyright   = $versionInfo.LegalCopyright
        Architecture     = $arch
        SignatureStatus  = if ($sig) { [string]$sig.Status } else { 'Unknown' }
        SignerSubject    = if ($sig -and $sig.SignerCertificate) { $sig.SignerCertificate.Subject } else { '' }
        SignerIssuer     = if ($sig -and $sig.SignerCertificate) { $sig.SignerCertificate.Issuer } else { '' }
        SignerThumbprint = if ($sig -and $sig.SignerCertificate) { $sig.SignerCertificate.Thumbprint } else { '' }
    }
}

function Test-ZipEntryExists {
    <#
    .SYNOPSIS
        Returns $true if the archive at $Path contains an entry matching $EntryName (exact) or $Pattern (wildcard).
    .DESCRIPTION
        Opens the ZIP read-only and checks its entries. Returns $false if the file is not a valid ZIP
        or does not contain the target. Safe to call on any file path.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Predicate against multiple zip entries; plural reads correctly.')]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$EntryName,
        [string]$Pattern,
        [switch]$RootOnly
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        foreach ($entry in $archive.Entries) {
            $full = $entry.FullName
            if ($RootOnly -and ($full -match '/' -or $full -match '\\')) { continue }
            if ($EntryName -and $full -eq $EntryName) { return $true }
            if ($Pattern -and ($entry.Name -like $Pattern)) { return $true }
        }
        return $false
    }
    catch {
        return $false
    }
    finally {
        if ($archive) { $archive.Dispose() }
    }
}

function Get-ZipEntryText {
    <#
    .SYNOPSIS
        Reads a text entry out of a ZIP archive and returns its contents as a string.
    .DESCRIPTION
        Returns $null if the archive cannot be opened or the entry does not exist.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$EntryName
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    $archive = $null; $stream = $null; $reader = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $entry = $archive.GetEntry($EntryName)
        if ($null -eq $entry) { return $null }
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream)
        return $reader.ReadToEnd()
    }
    catch {
        return $null
    }
    finally {
        if ($reader) { $reader.Dispose() }
        if ($stream) { $stream.Dispose() }
        if ($archive) { $archive.Dispose() }
    }
}

function Get-ZipRootEntryByPattern {
    <#
    .SYNOPSIS
        Returns the full path of the first ZIP entry at root matching a wildcard pattern, or $null.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Pattern
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -match '/' -or $entry.FullName -match '\\') { continue }
            if ($entry.Name -like $Pattern) { return $entry.FullName }
        }
        return $null
    }
    catch {
        return $null
    }
    finally {
        if ($archive) { $archive.Dispose() }
    }
}

function Test-IsZipFile {
    <#
    .SYNOPSIS
        Returns $true if the first 4 bytes of the file are the ZIP local-file-header magic (PK\x03\x04).
    #>
    param([Parameter(Mandatory)][string]$Path)

    $stream = $null
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        $bytes = New-Object byte[] 4
        $read = $stream.Read($bytes, 0, 4)
        if ($read -lt 4) { return $false }
        return ($bytes[0] -eq 0x50 -and $bytes[1] -eq 0x4B -and $bytes[2] -eq 0x03 -and $bytes[3] -eq 0x04)
    }
    catch {
        return $false
    }
    finally {
        if ($stream) { try { $stream.Close() } catch { $null = $_ } }
    }
}

function script:Find-SquirrelNupkgRefs {
    <#
    .SYNOPSIS
        Finds Squirrel-style "<AppId>-<Version>-(full|delta).nupkg" refs in a
        decoded byte buffer, case-insensitively.
    .OUTPUTS
        Array of PSCustomObject (FileName, AppName, Version, Kind). May be empty.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns an array of references; singular would suggest a single hit only.')]
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return @() }
    $hits = @()
    # Prerelease class now allows hyphens (NuGet/SemVer permits "1.2.3-beta-1").
    # Anchored on $ so the regex can't drift past the window end.
    $pattern = '([A-Za-z0-9][A-Za-z0-9.-]*)-(\d+\.\d+\.\d+(?:\.\d+)?(?:-[A-Za-z0-9.-]+)?)-(full|delta)\.nupkg$'
    $regex = [regex]::new($pattern,
        [System.Text.RegularExpressions.RegexOptions]::CultureInvariant -bor
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase,
        [TimeSpan]::FromMilliseconds(100))
    $needles = @('-full.nupkg', '-delta.nupkg')
    $budgetMs = 2000
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($needle in $needles) {
        $offset = 0
        while ($offset -lt $Text.Length) {
            if ($sw.ElapsedMilliseconds -ge $budgetMs) { return ,$hits }
            $idx = $Text.IndexOf($needle, $offset, [System.StringComparison]::OrdinalIgnoreCase)
            if ($idx -lt 0) { break }
            $start = [Math]::Max(0, $idx - 256)
            $end   = $idx + $needle.Length
            $window = $Text.Substring($start, $end - $start)
            try {
                $m = $regex.Match($window)
                if ($m.Success) {
                    $hits += [PSCustomObject]@{
                        FileName = $m.Value
                        AppName  = $m.Groups[1].Value
                        Version  = $m.Groups[2].Value
                        Kind     = $m.Groups[3].Value
                    }
                }
            } catch [System.Text.RegularExpressions.RegexMatchTimeoutException] {
                # Window-level safety net. Skip this occurrence and continue.
            }
            $offset = $end
        }
    }
    return ,$hits
}

function Get-InstallerType {
    <#
    .SYNOPSIS
        Detects installer framework by scanning binary signatures and ZIP layout.
    .DESCRIPTION
        Returns one of:
          MSI, NSIS, InnoSetup, InstallShield, WixBurn, AdvancedInstaller,
          BitRock, 7zSFX, WinRarSFX,
          Chocolatey, NuGet, Intunewin, Msix, MsixBundle, PsadtV3, PsadtV4,
          Squirrel,
          Unknown.
    #>
    param([Parameter(Mandatory)][string]$Path)

    # MSI / MSP by extension. MSPs are first-class so the analyzer can read
    # MsiPatchMetadata for the post-patch DisplayName / DisplayVersion (the
    # detection-method target for any base-MSI+cumulative-MSP product, e.g.
    # Adobe Reader, Office, large enterprise vendors).
    if ($Path -match '\.msi$') { return 'MSI' }
    if ($Path -match '\.msp$') { return 'MSP' }

    # NuGet / Chocolatey packages: .nupkg = ZIP + root *.nuspec
    if ($Path -match '\.nupkg$') {
        if (Test-IsZipFile -Path $Path) {
            $nuspec = Get-ZipRootEntryByPattern -Path $Path -Pattern '*.nuspec'
            if ($nuspec) {
                # Chocolatey if it carries a chocolatey*.ps1 under tools/
                if (Test-ZipEntryExists -Path $Path -Pattern 'chocolatey*.ps1') { return 'Chocolatey' }
                return 'NuGet'
            }
        }
    }

    # Intune Win32 Content Prep Tool output: .intunewin = OPC/ZIP with Metadata/Detection.xml
    if ($Path -match '\.intunewin$') {
        if ((Test-IsZipFile -Path $Path) -and (Test-ZipEntryExists -Path $Path -EntryName 'IntuneWinPackage/Metadata/Detection.xml')) {
            return 'Intunewin'
        }
    }

    # MSIX / APPX bundles (check before single packages -- bundle extensions are distinct)
    if ($Path -match '\.(msixbundle|appxbundle)$') {
        if ((Test-IsZipFile -Path $Path) -and (Test-ZipEntryExists -Path $Path -EntryName 'AppxMetadata/AppxBundleManifest.xml')) {
            return 'MsixBundle'
        }
    }

    # MSIX / APPX single packages
    if ($Path -match '\.(msix|appx)$') {
        if ((Test-IsZipFile -Path $Path) -and (Test-ZipEntryExists -Path $Path -EntryName 'AppxManifest.xml')) {
            return 'Msix'
        }
    }

    # PSADT (PSAppDeployToolkit) wrapped as ZIP. v4 tested before v3 because the v4
    # module folder can coexist with a legacy Deploy-Application.ps1 in compatibility
    # layouts; we want to classify those as v4.
    if ($Path -match '\.zip$') {
        if (Test-IsZipFile -Path $Path) {
            if ((Test-ZipEntryExists -Path $Path -Pattern 'PSAppDeployToolkit.psm1') -or
                (Test-ZipEntryExists -Path $Path -Pattern 'PSAppDeployToolkit.psd1') -or
                (Test-ZipEntryExists -Path $Path -Pattern 'Invoke-AppDeployToolkit.ps1')) {
                return 'PsadtV4'
            }
            if ((Test-ZipEntryExists -Path $Path -Pattern 'Deploy-Application.ps1') -and
                (Test-ZipEntryExists -Path $Path -Pattern 'AppDeployToolkitMain.ps1')) {
                return 'PsadtV3'
            }
        }
    }

    Write-Log "Scanning binary signatures..."

    # 4MB scan window. Smaller windows miss markers in modern installers
    # (BitRock's marker often sits 2-3 MB in).
    $maxRead = 4MB
    $fileSize = (Get-Item -LiteralPath $Path).Length
    $readSize = [Math]::Min($maxRead, $fileSize)

    $stream = $null
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        $bytes = New-Object byte[] $readSize
        $stream.Read($bytes, 0, $readSize) | Out-Null
    }
    finally {
        if ($stream) { try { $stream.Close() } catch { $null = $_ } }
    }

    $asciiText = [System.Text.Encoding]::ASCII.GetString($bytes)
    # Wide-char (UTF-16LE) view of the same buffer. Modern Windows-native Squirrel
    # builds (GitHub Desktop, Slack, Discord, Teams) keep their marker strings as
    # wide chars and the ASCII view misses them entirely. We scan both views.
    $wideText  = [System.Text.Encoding]::Unicode.GetString($bytes)

    # OLE Compound Document (MSI in disguise, e.g., .exe wrapping MSI)
    if ($bytes.Length -ge 8 -and $bytes[0] -eq 0xD0 -and $bytes[1] -eq 0xCF -and $bytes[2] -eq 0x11 -and $bytes[3] -eq 0xE0) {
        return 'MSI'
    }

    # Squirrel / Electron Setup.exe -- check BEFORE NSIS. Squirrel's bootstrapper is
    # not NSIS-built so it should not carry NullsoftInst markers, but being earlier
    # in the chain is defensive against future false positives. Require 2+ distinct
    # Squirrel markers to suppress accidental hits on the literal string "Squirrel".
    $squirrelMarkers = @('SquirrelTemp', 'squirrel-install', 'squirrel-updated',
                         'squirrel-uninstall', 'squirrel-firstrun', 'squirrel-obsolete')
    $squirrelHits = 0
    foreach ($m in $squirrelMarkers) {
        $esc = [regex]::Escape($m)
        if ($asciiText -match $esc -or $wideText -match $esc) { $squirrelHits++ }
    }
    if ($squirrelHits -ge 2) {
        return 'Squirrel'
    }
    # Weaker form: a single lifecycle marker combined with Update.exe string is enough
    $hasUpdateExe = ($asciiText -match 'Update\.exe') -or ($wideText -match 'Update\.exe')
    if ($squirrelHits -ge 1 -and $hasUpdateExe) {
        return 'Squirrel'
    }
    # Squirrel-nupkg-reference fallback: a "<AppId>-<Version>-full.nupkg"
    # ref plus "Update.exe" in either encoding is enough to classify, even
    # when the lifecycle markers live past the scan window.
    if ($hasUpdateExe -and (
            (Find-SquirrelNupkgRefs -Text $asciiText).Count -gt 0 -or
            (Find-SquirrelNupkgRefs -Text $wideText).Count  -gt 0)) {
        return 'Squirrel'
    }

    # WiX Burn -- check first (specific marker, avoids false positives)
    if ($asciiText -match 'WixBundleManifest|\.wixburn') {
        return 'WixBurn'
    }

    # NSIS -- DEADBEEF marker (little-endian) + NullsoftInst
    if ($asciiText -match 'NullsoftInst|Nullsoft\.NSIS') {
        return 'NSIS'
    }
    # Also check for DEADBEEF magic in raw bytes
    for ($i = 0; $i -lt [Math]::Min($bytes.Length - 4, 64KB); $i++) {
        if ($bytes[$i] -eq 0xEF -and $bytes[$i+1] -eq 0xBE -and $bytes[$i+2] -eq 0xAD -and $bytes[$i+3] -eq 0xDE) {
            return 'NSIS'
        }
    }

    # Inno Setup
    if ($asciiText -match 'Inno Setup') {
        return 'InnoSetup'
    }

    # InstallShield
    if ($asciiText -match 'InstallShield') {
        return 'InstallShield'
    }

    # Advanced Installer
    if ($asciiText -match 'Advanced Installer') {
        return 'AdvancedInstaller'
    }

    # BitRock InstallBuilder -- used by PostgreSQL, Bitnami, JFrog Artifactory,
    # and other vendors that need a cross-platform installer framework. The
    # signature string "BitRock" lives in the bootstrap metadata a couple MB in.
    if ($asciiText -match 'BitRock') {
        return 'BitRock'
    }

    # 7-Zip SFX -- look for 7z magic bytes anywhere in the file
    for ($i = 0; $i -lt [Math]::Min($bytes.Length - 6, $readSize); $i++) {
        if ($bytes[$i] -eq 0x37 -and $bytes[$i+1] -eq 0x7A -and $bytes[$i+2] -eq 0xBC -and
            $bytes[$i+3] -eq 0xAF -and $bytes[$i+4] -eq 0x27 -and $bytes[$i+5] -eq 0x1C) {
            return '7zSFX'
        }
    }
    if ($asciiText -match '!@InstallEnd@!') {
        return '7zSFX'
    }

    # WinRAR SFX -- RAR magic: 52 61 72 21 1A 07
    for ($i = 0; $i -lt [Math]::Min($bytes.Length - 7, $readSize); $i++) {
        if ($bytes[$i] -eq 0x52 -and $bytes[$i+1] -eq 0x61 -and $bytes[$i+2] -eq 0x72 -and
            $bytes[$i+3] -eq 0x21 -and $bytes[$i+4] -eq 0x1A -and $bytes[$i+5] -eq 0x07) {
            return 'WinRarSFX'
        }
    }

    return 'Unknown'
}

# ---------------------------------------------------------------------------
# NuGet / Chocolatey
# ---------------------------------------------------------------------------

function Get-ChocolateyMetadata {
    <#
    .SYNOPSIS
        Parses a Chocolatey or plain NuGet .nupkg and returns its nuspec metadata.
    .DESCRIPTION
        Returns a PSCustomObject with standardized fields (DisplayName, DisplayVersion,
        Publisher, Architecture, ProductCodeOrEquivalent, SilentInstallCommand,
        SilentUninstallCommand) plus a raw Nuspec hashtable of every element found.
        Returns $null if the package cannot be parsed.

        Detects Chocolatey vs plain NuGet by presence of tools/chocolateyInstall.ps1.
        Nuspec namespace: http://schemas.microsoft.com/packaging/2010/07/nuspec.xsd
        Source: https://learn.microsoft.com/en-us/nuget/reference/nuspec
    #>
    param([Parameter(Mandatory)][string]$Path)

    Write-Log "Reading nuspec from: $Path"

    if (-not (Test-IsZipFile -Path $Path)) { return $null }

    $nuspecPath = Get-ZipRootEntryByPattern -Path $Path -Pattern '*.nuspec'
    if (-not $nuspecPath) {
        Write-Log "No .nuspec at archive root" -Level WARN
        return $null
    }

    $xmlText = Get-ZipEntryText -Path $Path -EntryName $nuspecPath
    if (-not $xmlText) { return $null }

    # Strip BOM if present (some packagers emit UTF-8 BOM in nuspec)
    if ($xmlText.Length -gt 0 -and $xmlText[0] -eq [char]0xFEFF) {
        $xmlText = $xmlText.Substring(1)
    }

    try {
        $xml = [xml]$xmlText
    }
    catch {
        Write-Log "Failed to parse nuspec XML: $_" -Level ERROR
        return $null
    }

    # nuspec ships with multiple XSD namespace versions over the years:
    #   2010/07 - original NuGet spec
    #   2011/08 - adds licenseUrl/projectUrl extras
    #   2012/06 - adds developmentDependency
    #   2013/05 - current Chocolatey gallery default (added contentFiles, repo metadata)
    #   2017/09 - modern NuGet with license/repository/contentFiles
    # Any of these can show up in the wild. We register every known namespace
    # alias and probe each XPath in order, then fall back to no-namespace lookup.
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $aliases = @{
        'nu10' = 'http://schemas.microsoft.com/packaging/2010/07/nuspec.xsd'
        'nu11' = 'http://schemas.microsoft.com/packaging/2011/08/nuspec.xsd'
        'nu12' = 'http://schemas.microsoft.com/packaging/2012/06/nuspec.xsd'
        'nu13' = 'http://schemas.microsoft.com/packaging/2013/05/nuspec.xsd'
        'nu17' = 'http://schemas.microsoft.com/packaging/2017/09/nuspec.xsd'
    }
    foreach ($a in $aliases.GetEnumerator()) { $ns.AddNamespace($a.Key, $a.Value) }

    $meta = $null
    foreach ($alias in $aliases.Keys) {
        $meta = $xml.SelectSingleNode("/${alias}:package/${alias}:metadata", $ns)
        if ($meta) { break }
    }
    if (-not $meta) { $meta = $xml.SelectSingleNode('/package/metadata') }
    if (-not $meta) {
        Write-Log "No <metadata> element in nuspec" -Level WARN
        return $null
    }

    $nuspec = [ordered]@{}
    foreach ($child in $meta.ChildNodes) {
        if ($child.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            $nuspec[$child.LocalName] = [string]$child.InnerText
        }
    }

    $isChoco = Test-ZipEntryExists -Path $Path -Pattern 'chocolatey*.ps1'
    $installerType = if ($isChoco) { 'Chocolatey' } else { 'NuGet' }

    $displayName = if ($nuspec['title']) { $nuspec['title'] } else { $nuspec['id'] }
    $publisher   = if ($nuspec['authors']) { $nuspec['authors'] } else { $nuspec['owners'] }

    $silentInstall = if ($isChoco) {
        "choco install $($nuspec['id']) --version=$($nuspec['version']) -y --source=`"<SourceDirOrFeed>`""
    } else {
        "nuget install $($nuspec['id']) -Version $($nuspec['version']) -Source `"<SourceDirOrFeed>`""
    }
    $silentUninstall = if ($isChoco) {
        "choco uninstall $($nuspec['id']) -y"
    } else {
        'N/A (NuGet is a package source, not an installer)'
    }

    return [PSCustomObject]@{
        InstallerType            = $installerType
        PackageId                = $nuspec['id']
        DisplayName              = $displayName
        DisplayVersion           = $nuspec['version']
        Publisher                = $publisher
        Architecture             = 'N/A (package manifest)'
        ProductCodeOrEquivalent  = $nuspec['id']
        # Top-level conveniences. The Overview renderer reads $pkg.ProjectUrl /
        # Tags / Description and would otherwise see them as missing -- the
        # nuspec keys ($pkg.Nuspec.projectUrl, etc.) use lowercase camelCase
        # while the renderer expects PascalCase top-level properties. Bridge here.
        ProjectUrl               = $nuspec['projectUrl']
        Tags                     = $nuspec['tags']
        Description              = $nuspec['description']
        LicenseUrl               = $nuspec['licenseUrl']
        SilentInstallCommand     = $silentInstall
        SilentUninstallCommand   = $silentUninstall
        IsChocolatey             = $isChoco
        Nuspec                   = $nuspec
    }
}

# ---------------------------------------------------------------------------
# Intune Win32 (.intunewin)
# ---------------------------------------------------------------------------

function Get-IntunewinMetadata {
    <#
    .SYNOPSIS
        Parses a .intunewin package and returns its ApplicationInfo / Detection.xml metadata.
    .DESCRIPTION
        .intunewin is an OPC (ZIP) package produced by the Microsoft Win32 Content Prep Tool.
        The real metadata lives at IntuneWinPackage/Metadata/Detection.xml. The inner
        IntunePackage.intunewin is AES-encrypted and can only be decrypted by Intune itself
        using keys from the service; this function does NOT attempt decryption.

        Source: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool
                https://svrooij.io/2023/10/04/analysing-win32-content-prep-tool/

        By default, AES key material (EncryptionKey, MacKey, InitializationVector, Mac)
        is redacted in the returned EncryptionInfo block; the Encrypted boolean and the
        non-secret fields (ProfileIdentifier, FileDigest, FileDigestAlgorithm) are still
        surfaced. Pass -IncludeIntunewinKeyMaterial to surface the raw values verbatim
        from Detection.xml.

        Returns $null if the package cannot be parsed.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$IncludeIntunewinKeyMaterial
    )

    Write-Log "Reading Intunewin Detection.xml from: $Path"

    if (-not (Test-IsZipFile -Path $Path)) { return $null }

    $xmlText = Get-ZipEntryText -Path $Path -EntryName 'IntuneWinPackage/Metadata/Detection.xml'
    if (-not $xmlText) {
        Write-Log "No IntuneWinPackage/Metadata/Detection.xml entry" -Level WARN
        return $null
    }

    if ($xmlText.Length -gt 0 -and $xmlText[0] -eq [char]0xFEFF) {
        $xmlText = $xmlText.Substring(1)
    }

    try {
        $xml = [xml]$xmlText
    }
    catch {
        Write-Log "Failed to parse Detection.xml: $_" -Level ERROR
        return $null
    }

    $appInfo = $xml.ApplicationInfo
    if (-not $appInfo) {
        Write-Log "Detection.xml missing <ApplicationInfo> root" -Level WARN
        return $null
    }

    $toolVersion = $appInfo.ToolVersion
    $name        = [string]$appInfo.Name
    $fileName    = [string]$appInfo.FileName
    $setupFile   = [string]$appInfo.SetupFile
    $contentSize = [string]$appInfo.UnencryptedContentSize

    $encInfo = $appInfo.EncryptionInfo
    $enc = if ($encInfo) {
        if ($IncludeIntunewinKeyMaterial) {
            [ordered]@{
                Encrypted            = $true
                EncryptionKey        = [string]$encInfo.EncryptionKey
                MacKey               = [string]$encInfo.MacKey
                InitializationVector = [string]$encInfo.InitializationVector
                Mac                  = [string]$encInfo.Mac
                ProfileIdentifier    = [string]$encInfo.ProfileIdentifier
                FileDigest           = [string]$encInfo.FileDigest
                FileDigestAlgorithm  = [string]$encInfo.FileDigestAlgorithm
            }
        } else {
            [ordered]@{
                Encrypted            = $true
                EncryptionKey        = '<redacted>'
                MacKey               = '<redacted>'
                InitializationVector = '<redacted>'
                Mac                  = '<redacted>'
                ProfileIdentifier    = [string]$encInfo.ProfileIdentifier
                FileDigest           = [string]$encInfo.FileDigest
                FileDigestAlgorithm  = [string]$encInfo.FileDigestAlgorithm
            }
        }
    } else { [ordered]@{} }

    $msiSource = $false
    $msi = $null
    if ($appInfo.MsiInfo) {
        $msiSource = $true
        $mi = $appInfo.MsiInfo
        $msi = [ordered]@{
            MsiProductCode      = [string]$mi.MsiProductCode
            MsiProductVersion   = [string]$mi.MsiProductVersion
            MsiUpgradeCode      = [string]$mi.MsiUpgradeCode
            MsiExecutionContext = [string]$mi.MsiExecutionContext
            MsiRequiresLogon    = [string]$mi.MsiRequiresLogon
            MsiRequiresReboot   = [string]$mi.MsiRequiresReboot
            MsiIsMachineInstall = [string]$mi.MsiIsMachineInstall
            MsiIsUserInstall    = [string]$mi.MsiIsUserInstall
            MsiPackageCode      = [string]$mi.MsiPackageCode
            MsiPublisher        = [string]$mi.MsiPublisher
        }
    }

    $displayName = if ($name) { $name } else { $setupFile }
    $displayVersion = if ($msi) { $msi['MsiProductVersion'] } else { '' }
    $publisher = if ($msi) { $msi['MsiPublisher'] } else { '' }
    $productCode = if ($msi) { $msi['MsiProductCode'] } else { $fileName }

    # Architecture: MSI context hints at System (per-machine) vs User (per-user) but not x86/x64.
    # The Intune package itself is architecture-agnostic; the original setup file carried the arch.
    $architecture = if ($msiSource) {
        switch ($msi['MsiExecutionContext']) {
            'System' { 'Per-machine (MSI)' }
            'User'   { 'Per-user (MSI)' }
            'Any'    { 'Per-machine or per-user (MSI)' }
            default  { 'N/A (see embedded setup file)' }
        }
    } else {
        'N/A (see embedded setup file)'
    }

    $silentInstall = if ($msiSource) {
        "msiexec.exe /i `"<ExtractedSetup>`" /qn /norestart"
    } else {
        "`"<ExtractedSetup>`" <OriginalSilentSwitches>"
    }
    $silentUninstall = if ($msiSource -and $msi['MsiProductCode']) {
        "msiexec.exe /x `"$($msi['MsiProductCode'])`" /qn /norestart"
    } else {
        'N/A (uninstall command defined in Intune portal, not in the .intunewin)'
    }

    return [PSCustomObject]@{
        InstallerType           = 'Intunewin'
        ToolVersion             = $toolVersion
        Name                    = $name
        DisplayName             = $displayName
        DisplayVersion          = $displayVersion
        Publisher               = $publisher
        Architecture            = $architecture
        ProductCodeOrEquivalent = $productCode
        SetupFile               = $setupFile
        FileName                = $fileName
        UnencryptedContentSize  = $contentSize
        IsMsiSource             = $msiSource
        MsiInfo                 = $msi
        EncryptionInfo          = $enc
        SilentInstallCommand    = $silentInstall
        SilentUninstallCommand  = $silentUninstall
    }
}

# ---------------------------------------------------------------------------
# MSIX / APPX (single package + bundle)
# ---------------------------------------------------------------------------

function Get-MsixManifest {
    <#
    .SYNOPSIS
        Parses an MSIX / APPX package (single or bundle) and returns its identity metadata.
    .DESCRIPTION
        Reads AppxManifest.xml for single packages (namespace
        http://schemas.microsoft.com/appx/manifest/foundation/windows10) or
        AppxMetadata/AppxBundleManifest.xml for bundles (namespace
        http://schemas.microsoft.com/appx/2013/bundle). Returns a standardized
        PSCustomObject plus raw Identity / Packages data.

        Returns $null if neither manifest is present or parseable.
    #>
    param([Parameter(Mandatory)][string]$Path)

    Write-Log "Reading MSIX/APPX manifest from: $Path"

    if (-not (Test-IsZipFile -Path $Path)) { return $null }

    # Try bundle first: AppxMetadata/AppxBundleManifest.xml
    $bundleText = Get-ZipEntryText -Path $Path -EntryName 'AppxMetadata/AppxBundleManifest.xml'
    if ($bundleText) {
        return (ConvertFrom-MsixBundleManifest -XmlText $bundleText)
    }

    # Fall through to single package: AppxManifest.xml
    $manifestText = Get-ZipEntryText -Path $Path -EntryName 'AppxManifest.xml'
    if ($manifestText) {
        return (ConvertFrom-MsixPackageManifest -XmlText $manifestText)
    }

    Write-Log "No AppxManifest.xml or AppxBundleManifest.xml entry found" -Level WARN
    return $null
}

function ConvertFrom-MsixPackageManifest {
    <#
    .SYNOPSIS
        Parses the raw XML text of an AppxManifest.xml from a single-package .msix/.appx.
    #>
    param([Parameter(Mandatory)][string]$XmlText)

    if ($XmlText.Length -gt 0 -and $XmlText[0] -eq [char]0xFEFF) {
        $XmlText = $XmlText.Substring(1)
    }

    try { $xml = [xml]$XmlText }
    catch {
        Write-Log "Failed to parse AppxManifest.xml: $_" -Level ERROR
        return $null
    }

    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('p',    'http://schemas.microsoft.com/appx/manifest/foundation/windows10')
    $ns.AddNamespace('uap',  'http://schemas.microsoft.com/appx/manifest/uap/windows10')
    $ns.AddNamespace('rescap', 'http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities')

    $pkg = $xml.SelectSingleNode('/p:Package', $ns)
    if (-not $pkg) { $pkg = $xml.SelectSingleNode('/Package') }
    if (-not $pkg) {
        Write-Log "AppxManifest.xml has no <Package> root" -Level WARN
        return $null
    }

    $identity = $pkg.SelectSingleNode('p:Identity', $ns)
    if (-not $identity) { $identity = $pkg.SelectSingleNode('Identity') }
    $properties = $pkg.SelectSingleNode('p:Properties', $ns)
    if (-not $properties) { $properties = $pkg.SelectSingleNode('Properties') }

    $id = if ($identity) {
        [ordered]@{
            Name                 = [string]$identity.Name
            Publisher            = [string]$identity.Publisher
            Version              = [string]$identity.Version
            ProcessorArchitecture = [string]$identity.ProcessorArchitecture
            ResourceId           = [string]$identity.ResourceId
        }
    } else { [ordered]@{} }

    $displayName = ''; $publisherDisplay = ''; $description = ''; $logo = ''
    if ($properties) {
        foreach ($child in $properties.ChildNodes) {
            if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            switch ($child.LocalName) {
                'DisplayName'          { $displayName = [string]$child.InnerText }
                'PublisherDisplayName' { $publisherDisplay = [string]$child.InnerText }
                'Description'          { $description = [string]$child.InnerText }
                'Logo'                 { $logo = [string]$child.InnerText }
            }
        }
    }

    $architecture = if ($id['ProcessorArchitecture']) { $id['ProcessorArchitecture'] } else { 'Neutral' }

    return [PSCustomObject]@{
        InstallerType           = 'Msix'
        PackageKind             = 'SinglePackage'
        DisplayName             = if ($displayName) { $displayName } else { $id['Name'] }
        DisplayVersion          = $id['Version']
        Publisher               = if ($publisherDisplay) { $publisherDisplay } else { $id['Publisher'] }
        Architecture            = $architecture
        ProductCodeOrEquivalent = $id['Name']
        Identity                = $id
        PropertiesDescription   = $description
        PropertiesLogo          = $logo
        BundledPackages         = @()
        SilentInstallCommand    = 'Add-AppxPackage -Path "<msix>"'
        SilentUninstallCommand  = 'Remove-AppxPackage -Package "<PackageFullName>"'
    }
}

function ConvertFrom-MsixBundleManifest {
    <#
    .SYNOPSIS
        Parses the raw XML text of an AppxBundleManifest.xml from a .msixbundle / .appxbundle.
    #>
    param([Parameter(Mandatory)][string]$XmlText)

    if ($XmlText.Length -gt 0 -and $XmlText[0] -eq [char]0xFEFF) {
        $XmlText = $XmlText.Substring(1)
    }

    try { $xml = [xml]$XmlText }
    catch {
        Write-Log "Failed to parse AppxBundleManifest.xml: $_" -Level ERROR
        return $null
    }

    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('b', 'http://schemas.microsoft.com/appx/2013/bundle')

    $bundle = $xml.SelectSingleNode('/b:Bundle', $ns)
    if (-not $bundle) { $bundle = $xml.SelectSingleNode('/Bundle') }
    if (-not $bundle) {
        Write-Log "AppxBundleManifest.xml has no <Bundle> root" -Level WARN
        return $null
    }

    $identity = $bundle.SelectSingleNode('b:Identity', $ns)
    if (-not $identity) { $identity = $bundle.SelectSingleNode('Identity') }

    $id = if ($identity) {
        [ordered]@{
            Name      = [string]$identity.Name
            Publisher = [string]$identity.Publisher
            Version   = [string]$identity.Version
        }
    } else { [ordered]@{} }

    $pkgNodes = $bundle.SelectNodes('b:Packages/b:Package', $ns)
    if (-not $pkgNodes -or $pkgNodes.Count -eq 0) {
        $pkgNodes = $bundle.SelectNodes('Packages/Package')
    }
    $bundled = @()
    if ($pkgNodes) {
        foreach ($p in $pkgNodes) {
            $bundled += [PSCustomObject]@{
                Type         = [string]$p.Type
                Version      = [string]$p.Version
                Architecture = [string]$p.Architecture
                FileName     = [string]$p.FileName
                ResourceId   = [string]$p.ResourceId
            }
        }
    }

    $architectures = @($bundled | Where-Object { $_.Architecture } | ForEach-Object Architecture | Sort-Object -Unique) -join ', '
    if (-not $architectures) { $architectures = 'Bundle' }

    return [PSCustomObject]@{
        InstallerType           = 'MsixBundle'
        PackageKind             = 'Bundle'
        DisplayName             = $id['Name']
        DisplayVersion          = $id['Version']
        Publisher               = $id['Publisher']
        Architecture            = $architectures
        ProductCodeOrEquivalent = $id['Name']
        Identity                = $id
        PropertiesDescription   = ''
        PropertiesLogo          = ''
        BundledPackages         = $bundled
        SilentInstallCommand    = 'Add-AppxPackage -Path "<msixbundle>"'
        SilentUninstallCommand  = 'Remove-AppxPackage -Package "<PackageFullName>"'
    }
}

# ---------------------------------------------------------------------------
# PSAppDeployToolkit (PSADT)
# ---------------------------------------------------------------------------

function Get-ZipEntryPathByPattern {
    <#
    .SYNOPSIS
        Returns the first ZIP entry full path whose file-name matches a wildcard, or $null.
    .DESCRIPTION
        Searches the whole archive (any depth). Useful for sentinel files that may live
        under a top-level folder (e.g. PSADT's AppDeployToolkit/AppDeployToolkitMain.ps1).
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Pattern
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        foreach ($entry in $archive.Entries) {
            if ($entry.Name -like $Pattern) { return $entry.FullName }
        }
        return $null
    }
    catch {
        return $null
    }
    finally {
        if ($archive) { $archive.Dispose() }
    }
}

function ConvertFrom-PsadtDeployApplication {
    <#
    .SYNOPSIS
        Parses Deploy-Application.ps1 / Invoke-AppDeployToolkit.ps1 header variables
        ($appVendor, $appName, $appVersion, $appArch, $appLang, $appRevision,
         $appScriptVersion, $appScriptDate, $appScriptAuthor).
    .DESCRIPTION
        Works on both v3 (`[String]$appName = 'X'`) and v4 (`$adtSession = @{ AppName = 'X' }`)
        layouts. Returns an ordered hashtable of whichever fields were found.
    #>
    param([Parameter(Mandatory)][string]$ScriptText)

    $result = [ordered]@{
        AppVendor      = ''
        AppName        = ''
        AppVersion     = ''
        AppArch        = ''
        AppLang        = ''
        AppRevision    = ''
        ScriptVersion  = ''
        ScriptDate     = ''
        ScriptAuthor   = ''
    }

    # v3 pattern: [String]$appName = 'X'  or  [String]$appName = "X"
    $v3Patterns = [ordered]@{
        AppVendor     = '(?im)^\s*\[String\]\s*\$appVendor\s*=\s*[''"]([^''"]*)[''"]'
        AppName       = '(?im)^\s*\[String\]\s*\$appName\s*=\s*[''"]([^''"]*)[''"]'
        AppVersion    = '(?im)^\s*\[String\]\s*\$appVersion\s*=\s*[''"]([^''"]*)[''"]'
        AppArch       = '(?im)^\s*\[String\]\s*\$appArch\s*=\s*[''"]([^''"]*)[''"]'
        AppLang       = '(?im)^\s*\[String\]\s*\$appLang\s*=\s*[''"]([^''"]*)[''"]'
        AppRevision   = '(?im)^\s*\[String\]\s*\$appRevision\s*=\s*[''"]([^''"]*)[''"]'
        ScriptVersion = '(?im)^\s*\[String\]\s*\$appScriptVersion\s*=\s*[''"]([^''"]*)[''"]'
        ScriptDate    = '(?im)^\s*\[String\]\s*\$appScriptDate\s*=\s*[''"]([^''"]*)[''"]'
        ScriptAuthor  = '(?im)^\s*\[String\]\s*\$appScriptAuthor\s*=\s*[''"]([^''"]*)[''"]'
    }
    foreach ($key in $v3Patterns.Keys) {
        $m = [regex]::Match($ScriptText, $v3Patterns[$key])
        if ($m.Success) { $result[$key] = $m.Groups[1].Value }
    }

    # v4 pattern: key inside @{ AppName = 'X' } (loose match; same InnerText target)
    $v4Patterns = [ordered]@{
        AppVendor     = '(?im)\bAppVendor\s*=\s*[''"]([^''"]*)[''"]'
        AppName       = '(?im)\bAppName\s*=\s*[''"]([^''"]*)[''"]'
        AppVersion    = '(?im)\bAppVersion\s*=\s*[''"]([^''"]*)[''"]'
        AppArch       = '(?im)\bAppArch\s*=\s*[''"]([^''"]*)[''"]'
        AppLang       = '(?im)\bAppLang\s*=\s*[''"]([^''"]*)[''"]'
        AppRevision   = '(?im)\bAppRevision\s*=\s*[''"]([^''"]*)[''"]'
        ScriptVersion = '(?im)\bAppScriptVersion\s*=\s*[''"]([^''"]*)[''"]'
        ScriptDate    = '(?im)\bAppScriptDate\s*=\s*[''"]([^''"]*)[''"]'
        ScriptAuthor  = '(?im)\bAppScriptAuthor\s*=\s*[''"]([^''"]*)[''"]'
    }
    foreach ($key in $v4Patterns.Keys) {
        if ($result[$key]) { continue }   # v3 match takes precedence
        $m = [regex]::Match($ScriptText, $v4Patterns[$key])
        if ($m.Success) { $result[$key] = $m.Groups[1].Value }
    }

    return $result
}

function Get-PsadtMetadata {
    <#
    .SYNOPSIS
        Parses a PSAppDeployToolkit-wrapped ZIP and returns its toolkit version + per-app header.
    .DESCRIPTION
        Detects v3 (Deploy-Application.ps1 + AppDeployToolkit/AppDeployToolkitMain.ps1) or
        v4 (Invoke-AppDeployToolkit.ps1 + PSAppDeployToolkit module) layouts. Reads the
        toolkit engine version from the engine script/module and the per-app metadata
        ($appName, $appVersion, $appVendor, etc.) from the deployment script.

        Sources:
          https://psappdeploytoolkit.com/docs/4.0.x/deployment-concepts/deployment-structure
          https://psappdeploytoolkit.com/docs/deployment-concepts/invoke-appdeploytoolkit
    #>
    param([Parameter(Mandatory)][string]$Path)

    Write-Log "Reading PSADT layout from: $Path"

    if (-not (Test-IsZipFile -Path $Path)) { return $null }

    # v4 detection: module manifest OR module psm1 OR Invoke-AppDeployToolkit.ps1
    $v4ManifestPath = Get-ZipEntryPathByPattern -Path $Path -Pattern 'PSAppDeployToolkit.psd1'
    $v4ModulePath   = Get-ZipEntryPathByPattern -Path $Path -Pattern 'PSAppDeployToolkit.psm1'
    $v4InvokePath   = Get-ZipEntryPathByPattern -Path $Path -Pattern 'Invoke-AppDeployToolkit.ps1'

    # v3 detection: AppDeployToolkitMain.ps1 + Deploy-Application.ps1
    $v3MainPath     = Get-ZipEntryPathByPattern -Path $Path -Pattern 'AppDeployToolkitMain.ps1'
    $v3DeployPath   = Get-ZipEntryPathByPattern -Path $Path -Pattern 'Deploy-Application.ps1'

    $isV4 = [bool]($v4ManifestPath -or $v4ModulePath -or $v4InvokePath)
    $isV3 = [bool]($v3MainPath -and $v3DeployPath)

    if (-not $isV4 -and -not $isV3) {
        Write-Log "No PSADT sentinel files found in archive" -Level WARN
        return $null
    }

    $toolkitVersion = ''
    $toolkitVariant = if ($isV4) { 'v4' } else { 'v3' }

    if ($isV4) {
        if ($v4ManifestPath) {
            $psd1Text = Get-ZipEntryText -Path $Path -EntryName $v4ManifestPath
            if ($psd1Text) {
                $m = [regex]::Match($psd1Text, "(?im)^\s*ModuleVersion\s*=\s*['""]([^'""]+)['""]")
                if ($m.Success) { $toolkitVersion = $m.Groups[1].Value }
            }
        }
    }
    else {
        if ($v3MainPath) {
            $mainText = Get-ZipEntryText -Path $Path -EntryName $v3MainPath
            if ($mainText) {
                # Pattern: [Version]$appDeployMainScriptVersion = [Version]'3.9.2'
                $m = [regex]::Match($mainText, "(?im)\`$appDeployMainScriptVersion\s*=\s*\[Version\]\s*['""]([^'""]+)['""]")
                if ($m.Success) {
                    $toolkitVersion = $m.Groups[1].Value
                } else {
                    $m2 = [regex]::Match($mainText, "(?im)\`$appDeployMainScriptVersion\s*=\s*['""]([^'""]+)['""]")
                    if ($m2.Success) { $toolkitVersion = $m2.Groups[1].Value }
                }
            }
        }
    }

    # Per-app metadata: prefer Invoke-AppDeployToolkit.ps1 on v4, Deploy-Application.ps1 on v3
    $appScriptPath = if ($isV4 -and $v4InvokePath) { $v4InvokePath } elseif ($v3DeployPath) { $v3DeployPath } else { $null }
    $appMeta = [ordered]@{}
    if ($appScriptPath) {
        $scriptText = Get-ZipEntryText -Path $Path -EntryName $appScriptPath
        if ($scriptText) {
            $appMeta = ConvertFrom-PsadtDeployApplication -ScriptText $scriptText
        }
    }

    $displayName = if ($appMeta['AppName']) {
        if ($appMeta['AppVendor']) { "$($appMeta['AppVendor']) $($appMeta['AppName'])" } else { $appMeta['AppName'] }
    } else { '' }

    $architecture = if ($appMeta['AppArch']) { $appMeta['AppArch'] } else { 'N/A (see embedded installer)' }
    $installerType = if ($isV4) { 'PsadtV4' } else { 'PsadtV3' }

    $silentInstall = if ($isV4) {
        'Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent'
    } else {
        'Deploy-Application.exe -DeploymentType Install -DeployMode Silent'
    }
    $silentUninstall = if ($isV4) {
        'Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent'
    } else {
        'Deploy-Application.exe -DeploymentType Uninstall -DeployMode Silent'
    }

    return [PSCustomObject]@{
        InstallerType           = $installerType
        ToolkitVariant          = $toolkitVariant
        ToolkitVersion          = $toolkitVersion
        DisplayName             = $displayName
        DisplayVersion          = $appMeta['AppVersion']
        Publisher               = $appMeta['AppVendor']
        Architecture            = $architecture
        ProductCodeOrEquivalent = if ($appMeta['AppName']) { $appMeta['AppName'] } else { 'N/A' }
        AppMetadata             = $appMeta
        EngineScriptPath        = if ($isV4) { $v4ManifestPath } else { $v3MainPath }
        DeploymentScriptPath    = $appScriptPath
        SilentInstallCommand    = $silentInstall
        SilentUninstallCommand  = $silentUninstall
    }
}

# ---------------------------------------------------------------------------
# Squirrel / Electron Setup.exe
# ---------------------------------------------------------------------------

function Get-SquirrelMetadata {
    <#
    .SYNOPSIS
        Extracts Squirrel / Electron Setup.exe metadata from embedded refs.
    .DESCRIPTION
        Binary-string scans the first 4MB for AppName/Version (from
        "<Name>-<Version>-full.nupkg"), update-feed URLs, and Squirrel
        lifecycle markers. Returns a PackageMetadata PSCustomObject; $null
        when no markers are found.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$MaxBytes = 4MB
    )

    Write-Log "Scanning Squirrel binary: $Path"

    if (-not (Test-Path -LiteralPath $Path)) { return $null }

    $stream = $null
    try {
        $fileSize = (Get-Item -LiteralPath $Path).Length
        $readSize = [Math]::Min($MaxBytes, $fileSize)
        $stream = [System.IO.File]::OpenRead($Path)
        $bytes = New-Object byte[] $readSize
        [void]$stream.Read($bytes, 0, $readSize)
    }
    catch {
        Write-Log "Failed to read binary: $_" -Level ERROR
        return $null
    }
    finally {
        if ($stream) { try { $stream.Close() } catch { $null = $_ } }
    }

    $asciiText = [System.Text.Encoding]::ASCII.GetString($bytes)
    # Wide-char (UTF-16LE) view of the same buffer. Modern Windows-native Squirrel
    # builds keep markers as wide chars (GitHub Desktop, Slack, Discord, Teams).
    $wideText  = [System.Text.Encoding]::Unicode.GetString($bytes)

    $markerPatterns = @(
        'SquirrelTemp', 'squirrel-install', 'squirrel-updated',
        'squirrel-uninstall', 'squirrel-firstrun', 'squirrel-obsolete'
    )
    $markersFound = @()
    foreach ($m in $markerPatterns) {
        $esc = [regex]::Escape($m)
        if ($asciiText -match $esc -or $wideText -match $esc) { $markersFound += $m }
    }
    $hasUpdateExe = ($asciiText -match 'Update\.exe') -or ($wideText -match 'Update\.exe')

    if ($markersFound.Count -eq 0 -and -not $hasUpdateExe) {
        Write-Log "No Squirrel markers detected in first $readSize bytes" -Level WARN
        return $null
    }

    # Extract embedded nupkg references:
    #   <AppId>-<Version>-full.nupkg
    #   <AppId>-<Version>-delta.nupkg
    $nupkgRefs = @()
    $nupkgRefs += (Find-SquirrelNupkgRefs -Text $asciiText)
    $nupkgRefs += (Find-SquirrelNupkgRefs -Text $wideText)
    # De-dup (an Electron build may have the same nupkg as both ASCII and UTF-16).
    $nupkgRefs = @($nupkgRefs | Sort-Object FileName -Unique)
    $appName = ''
    $version = ''
    if ($nupkgRefs.Count -gt 0) {
        $preferred = $nupkgRefs | Where-Object Kind -EQ 'full' | Select-Object -First 1
        if (-not $preferred) { $preferred = $nupkgRefs[0] }
        $appName = $preferred.AppName
        $version = $preferred.Version
    }

    # URLs (may include the update feed; these are hints, not authoritative)
    $urls = @(
        ([regex]::Matches($asciiText, 'https?://[^\s"''<>\x00]+') | ForEach-Object { $_.Value }) +
        ([regex]::Matches($wideText,  'https?://[^\s"''<>\x00]+') | ForEach-Object { $_.Value }) |
        Sort-Object -Unique |
        Select-Object -First 20
    )

    $confidence = if ($markersFound.Count -ge 3 -and $nupkgRefs.Count -gt 0) { 'High' }
                  elseif ($markersFound.Count -ge 2) { 'High' }
                  elseif ($markersFound.Count -ge 1 -and $hasUpdateExe) { 'Medium' }
                  elseif ($hasUpdateExe) { 'Low' }
                  else { 'Low' }

    $displayName = if ($appName) { $appName } else { [System.IO.Path]::GetFileNameWithoutExtension($Path) }

    return [PSCustomObject]@{
        InstallerType           = 'Squirrel'
        DisplayName             = $displayName
        DisplayVersion          = $version
        Publisher               = ''
        Architecture            = 'N/A (see embedded PE of bundled app)'
        ProductCodeOrEquivalent = $appName
        MarkersFound            = $markersFound
        HasUpdateExe            = $hasUpdateExe
        NupkgReferences         = $nupkgRefs
        ObservedUrls            = $urls
        Confidence              = $confidence
        SilentInstallCommand    = '"<Setup.exe>" --silent'
        SilentUninstallCommand  = '"%LOCALAPPDATA%\<AppName>\Update.exe" --uninstall -s'
    }
}

# ---------------------------------------------------------------------------
# WiX Burn bundle
# ---------------------------------------------------------------------------

function Get-WixBurnMetadata {
    <#
    .SYNOPSIS
        Extracts the BundleId GUID from a WiX Burn bundle's .wixburn PE section.
    .DESCRIPTION
        WiX Burn (v3+) carries a BURN_SECTION header at the start of the .wixburn
        section. The structure begins with magic 0x00f14300, a version DWORD, then
        the bundle GUID at offset 0x08. The bundle's ARP key under
        HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\ is named with
        this GUID (per-machine bundles) or under HKCU for per-user bundles.
        We can't distinguish per-machine vs per-user from the header alone -- that
        lives in the manifest XML inside the embedded CAB container -- so we
        default the predicted key to HKLM and note the alternative.
        Returns $null if the section is missing or unreadable.
        Source: WiX toolset Burn source, src\engine\section.h (BURN_SECTION).
    #>
    param([Parameter(Mandatory)][string]$Path)

    Write-Log "Reading WiX Burn .wixburn section: $Path"

    if (-not (Test-Path -LiteralPath $Path)) { return $null }

    $stream = $null
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        $br = New-Object System.IO.BinaryReader($stream)

        # DOS header -> e_lfanew at 0x3C
        [void]$stream.Seek(0x3C, 'Begin')
        $peOffset = $br.ReadUInt32()
        if ($peOffset -le 0 -or $peOffset -gt ($stream.Length - 24)) { return $null }

        # PE signature
        [void]$stream.Seek($peOffset, 'Begin')
        $sig = $br.ReadUInt32()
        if ($sig -ne 0x00004550) { return $null }    # 'PE\0\0'

        # COFF header
        [void]$br.ReadUInt16()                       # Machine
        $numSections = $br.ReadUInt16()
        [void]$br.ReadUInt32()                       # TimeDateStamp
        [void]$br.ReadUInt32()                       # PointerToSymbolTable
        [void]$br.ReadUInt32()                       # NumberOfSymbols
        $sizeOfOpt = $br.ReadUInt16()
        [void]$br.ReadUInt16()                       # Characteristics

        # Sanity-bound the section count. Real PE files run ~5-20 sections; a
        # crafted or corrupt header could claim up to 65535 here and burn the
        # log on 65k failed ReadBytes() calls inside the section walk. 128 is
        # well above any plausible real binary.
        if ($numSections -le 0 -or $numSections -gt 128) {
            Write-Log ("Implausible PE section count: $numSections") -Level WARN
            return $null
        }

        # Skip optional header to reach section table
        [void]$stream.Seek($sizeOfOpt, 'Current')

        $burnOffset = 0
        for ($i = 0; $i -lt $numSections; $i++) {
            $nameBytes = $br.ReadBytes(8)
            $name = [System.Text.Encoding]::ASCII.GetString($nameBytes).TrimEnd([char]0)
            [void]$br.ReadUInt32()                   # VirtualSize
            [void]$br.ReadUInt32()                   # VirtualAddress
            [void]$br.ReadUInt32()                   # SizeOfRawData
            $rawPtr = $br.ReadUInt32()
            [void]$br.ReadUInt32()                   # PointerToRelocations
            [void]$br.ReadUInt32()                   # PointerToLinenumbers
            [void]$br.ReadUInt16()                   # NumberOfRelocations
            [void]$br.ReadUInt16()                   # NumberOfLinenumbers
            [void]$br.ReadUInt32()                   # Characteristics
            if ($name -eq '.wixburn') { $burnOffset = $rawPtr; break }
        }

        if ($burnOffset -eq 0) {
            Write-Log "No .wixburn section in PE; not a WiX Burn bundle" -Level WARN
            return $null
        }

        # BURN_SECTION header: magic(4), version(4), guidBundleId(16), ...
        [void]$stream.Seek($burnOffset, 'Begin')
        $magic = $br.ReadUInt32()
        if ($magic -ne 0x00f14300) {
            Write-Log ("Unexpected .wixburn magic: 0x{0:X8}" -f $magic) -Level WARN
            return $null
        }
        $version = $br.ReadUInt32()
        $guidBytes = $br.ReadBytes(16)
        $bundleId = ([guid]::new($guidBytes)).ToString('B').ToUpperInvariant()

        $uninstallKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$bundleId"

        return [PSCustomObject]@{
            InstallerType        = 'WixBurn'
            BundleId             = $bundleId
            BurnVersion          = $version
            UninstallRegistryKey = $uninstallKey
            UninstallRegistryKeyNote = 'HKLM if per-machine; HKCU if per-user (per-machine flag lives in the embedded BundleManifest.xml).'
        }
    }
    catch {
        Write-Log "Failed to read WiX Burn header: $_" -Level WARN
        return $null
    }
    finally {
        if ($stream) { try { $stream.Dispose() } catch { $null = $_ } }
    }
}

# ---------------------------------------------------------------------------
# Package metadata dispatcher
# ---------------------------------------------------------------------------

function Get-PackageMetadataFor {
    <#
    .SYNOPSIS
        Dispatches to the matching Get-*Metadata extractor for the installer type.
    .DESCRIPTION
        Returns the PackageMetadata PSCustomObject for known framework types, or
        $null for types that don't have a framework-specific metadata extractor
        (MSI is handled via Get-MsiProperties / Get-MsiSummaryInfo elsewhere).
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Dispatches to a single Get-*Metadata extractor; returns one PackageMetadata object or null.')]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$InstallerType,
        [switch]$IncludeIntunewinKeyMaterial
    )
    switch ($InstallerType) {
        'Chocolatey' { return Get-ChocolateyMetadata -Path $Path }
        'NuGet'      { return Get-ChocolateyMetadata -Path $Path }
        'Intunewin'  { return Get-IntunewinMetadata  -Path $Path -IncludeIntunewinKeyMaterial:$IncludeIntunewinKeyMaterial }
        'Msix'       { return Get-MsixManifest       -Path $Path }
        'MsixBundle' { return Get-MsixManifest       -Path $Path }
        'PsadtV3'    { return Get-PsadtMetadata      -Path $Path }
        'PsadtV4'    { return Get-PsadtMetadata      -Path $Path }
        'Squirrel'   { return Get-SquirrelMetadata   -Path $Path }
        'WixBurn'    { return Get-WixBurnMetadata    -Path $Path }
        'MSP'        { return Get-MspMetadata        -Path $Path }
        default      { return $null }
    }
}

# ---------------------------------------------------------------------------
# Uninstall Registry Key resolution
# ---------------------------------------------------------------------------

function Get-UninstallRegistryKey {
    <#
    .SYNOPSIS
        Predicts the Add/Remove Programs registry key the installer writes to.
    .DESCRIPTION
        Returns a PSCustomObject with Path, Hive, Note. Returns $null when the
        installer type doesn't write to ARP.

        Per-type conventions:
          - MSI: HKLM\...\Uninstall\{ProductCode}; WOW6432Node for 32-bit MSIs on x64.
          - WiX Burn: BundleId from .wixburn header.
          - Squirrel: HKCU\...\Uninstall\<AppName>.
          - Inno Setup: HKLM\...\Uninstall\<AppId>_is1.
          - NSIS / InstallShield: HKLM\...\Uninstall\<script-defined>.
          - PSADT: HKLM\...\Uninstall\<AppName>.
          - MSIX: not in ARP; returns the Get-AppxPackage family-name path.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns one PSCustomObject with Path/Hive/Note fields; singular would imply a string return.')]
    param(
        [Parameter(Mandatory)][string]$InstallerType,
        [PSCustomObject]$FileInfo,
        [hashtable]$MsiProperties,
        [PSCustomObject]$PackageMetadata,
        [PSCustomObject]$DeploymentFields,
        # Optional override for MSI architecture detection. Get-InstallerFileInfo
        # returns 'N/A (see MSI Summary)' for .msi files because the Property table
        # has no x86/x64 signal; the Template comes from the SummaryInformation
        # stream. Pipeline callers pass Get-MsiSummaryInfo.Architecture here.
        [string]$MsiArchitecture = ''
    )

    $arpRoot64 = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall'
    $arpRoot32 = 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
    $arpRootHkcu = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall'

    switch ($InstallerType) {
        'MSI' {
            $productCode = ''
            if ($MsiProperties -and $MsiProperties.Contains('ProductCode')) { $productCode = [string]$MsiProperties['ProductCode'] }
            if (-not $productCode) { return $null }

            # Per-user MSI detection. Microsoft ALLUSERS / MSIINSTALLPERUSER semantics:
            #   ALLUSERS=1 (or '1')                           -> per-machine (HKLM)
            #   ALLUSERS=2 + MSIINSTALLPERUSER=1              -> per-user    (HKCU)
            #   ALLUSERS='' (empty / unset) + no per-user opt -> per-user    (HKCU)
            #   ALLUSERS not present -> conservatively per-machine
            $allUsers = if ($MsiProperties -and $MsiProperties.Contains('ALLUSERS')) { [string]$MsiProperties['ALLUSERS'] } else { $null }
            $perUserFlag = if ($MsiProperties -and $MsiProperties.Contains('MSIINSTALLPERUSER')) { [string]$MsiProperties['MSIINSTALLPERUSER'] } else { '' }
            $isPerUser = $false
            if ($null -ne $allUsers) {
                # ALLUSERS exists in the Property table.
                if ($allUsers -eq '') { $isPerUser = $true }                                   # bare per-user
                elseif ($allUsers -eq '2' -and $perUserFlag -eq '1') { $isPerUser = $true }     # dual-mode opting into per-user
            } else {
                # No ALLUSERS row: explicit MSIINSTALLPERUSER=1 still routes per-user.
                if ($perUserFlag -eq '1') { $isPerUser = $true }
            }

            if ($isPerUser) {
                return [PSCustomObject]@{
                    Path = "$arpRootHkcu\$productCode"
                    Hive = 'HKCU'
                    Note = 'MSI Property table indicates per-user install (ALLUSERS / MSIINSTALLPERUSER). ARP key is under HKCU, not HKLM.'
                }
            }

            # Per-machine. Architecture decides WOW6432Node redirect on x64 Windows.
            # Prefer the explicit MSI Template override (from Get-MsiSummaryInfo);
            # FileInfo.Architecture is 'N/A (see MSI Summary)' for .msi files.
            $arch = if ($MsiArchitecture) { $MsiArchitecture }
                    elseif ($FileInfo -and $FileInfo.Architecture) { [string]$FileInfo.Architecture }
                    else { '' }
            $is32Bit = $arch -match '^(x86|Intel|32)$'
            $path = if ($is32Bit) { "$arpRoot32\$productCode" } else { "$arpRoot64\$productCode" }
            $note = if ($is32Bit) {
                'MSI 32-bit on x64 Windows: WOW6432Node redirect applies. Per-user installs go under HKCU.'
            } elseif (-not $MsiArchitecture -and $FileInfo -and [string]$FileInfo.Architecture -eq 'N/A (see MSI Summary)') {
                'Architecture unknown (MSI Template not supplied). Assuming x64 / HKLM. Pass -MsiArchitecture from Get-MsiSummaryInfo for x86 MSIs to get the WOW6432Node path.'
            } else {
                'Per-machine MSI. Per-user installs (ALLUSERS=2 + MSIINSTALLPERUSER=1, or ALLUSERS empty) live under HKCU.'
            }
            return [PSCustomObject]@{ Path = $path; Hive = 'HKLM'; Note = $note }
        }

        'WixBurn' {
            if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['UninstallRegistryKey'] -and $PackageMetadata.UninstallRegistryKey) {
                $note = if ($PackageMetadata.PSObject.Properties['UninstallRegistryKeyNote']) { [string]$PackageMetadata.UninstallRegistryKeyNote } else { '' }
                return [PSCustomObject]@{ Path = [string]$PackageMetadata.UninstallRegistryKey; Hive = 'HKLM'; Note = $note }
            }
            # No PackageMetadata = no BundleId. Rather than render "{BundleId}"
            # as a literal in the path, drop the field; the user will know it's
            # genuinely missing instead of seeing a templated placeholder.
            return $null
        }

        'Squirrel' {
            $appName = ''
            if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['ProductCodeOrEquivalent']) { $appName = [string]$PackageMetadata.ProductCodeOrEquivalent }
            if (-not $appName -and $PackageMetadata -and $PackageMetadata.PSObject.Properties['DisplayName']) { $appName = [string]$PackageMetadata.DisplayName }
            # No usable AppName -> no predicted path. Better to omit than to
            # render an unsubstituted "<AppName>" template literal in the path.
            if (-not $appName) { return $null }
            return [PSCustomObject]@{
                Path = "$arpRootHkcu\$appName"
                Hive = 'HKCU'
                Note = ''
            }
        }

        'InnoSetup' {
            $appId = if ($DeploymentFields -and $DeploymentFields.DisplayName) { [string]$DeploymentFields.DisplayName } else { '' }
            if (-not $appId) { return $null }
            return [PSCustomObject]@{
                Path = "$arpRoot64\${appId}_is1"
                Hive = 'HKLM'
                Note = 'Inno Setup appends "_is1" to AppId. AppId is script-defined; often equals AppName. Per-user installs go under HKCU; 32-bit installs use WOW6432Node.'
            }
        }

        'NSIS' {
            # TeamViewer Host strips ALL FileVersionInfo from its NSIS bootstrapper,
            # so DisplayName resolves to '' and there's no way to predict the key
            # name. Returning $null is more honest than rendering "<DisplayName>"
            # as a literal -- the user can see the field is genuinely missing.
            $name = if ($DeploymentFields -and $DeploymentFields.DisplayName) { [string]$DeploymentFields.DisplayName } else { '' }
            if (-not $name) { return $null }
            return [PSCustomObject]@{
                Path = "$arpRoot64\$name"
                Hive = 'HKLM'
                Note = 'NSIS uninstall key name is script-defined (WriteUninstaller / WriteRegStr). Convention is DisplayName under HKLM; per-user installs go under HKCU.'
            }
        }

        'InstallShield' {
            $code = if ($MsiProperties -and $MsiProperties.Contains('ProductCode')) { [string]$MsiProperties['ProductCode'] } else { '' }
            if (-not $code) { return $null }
            return [PSCustomObject]@{
                Path = "$arpRoot64\$code"
                Hive = 'HKLM'
                Note = 'InstallShield variants: Basic MSI uses MSI ProductCode; InstallScript MSI same; pure InstallScript projects use a named key. 32-bit on x64 uses WOW6432Node.'
            }
        }

        { $_ -in 'PsadtV3','PsadtV4' } {
            $name = if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['AppMetadata'] -and $PackageMetadata.AppMetadata -and $PackageMetadata.AppMetadata['AppName']) {
                [string]$PackageMetadata.AppMetadata['AppName']
            } elseif ($DeploymentFields -and $DeploymentFields.DisplayName) { [string]$DeploymentFields.DisplayName } else { '' }
            if (-not $name) { return $null }
            return [PSCustomObject]@{
                Path = "$arpRoot64\$name"
                Hive = 'HKLM'
                Note = 'PSADT does not write ARP entries itself; the embedded installer does. This is the conventional path that PSADT detection rules key on (configure in the deployment script).'
            }
        }

        'BitRock' {
            $name = if ($DeploymentFields -and $DeploymentFields.DisplayName) { [string]$DeploymentFields.DisplayName } else { '' }
            if (-not $name) { return $null }
            return [PSCustomObject]@{
                Path = "$arpRoot64\$name"
                Hive = 'HKLM'
                Note = 'BitRock InstallBuilder writes a named key under HKLM by default; per-user installs go under HKCU.'
            }
        }

        'AdvancedInstaller' {
            $code = if ($MsiProperties -and $MsiProperties.Contains('ProductCode')) { [string]$MsiProperties['ProductCode'] } else { '' }
            if (-not $code) { return $null }
            return [PSCustomObject]@{
                Path = "$arpRoot64\$code"
                Hive = 'HKLM'
                Note = 'Advanced Installer EXE wrappers an MSI; ARP key is the wrapped MSI ProductCode.'
            }
        }

        { $_ -in 'Msix','MsixBundle' } {
            $famName = ''
            if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['Identity'] -and $PackageMetadata.Identity -and $PackageMetadata.Identity.Name) {
                $famName = [string]$PackageMetadata.Identity.Name
            }
            if (-not $famName) { return $null }
            return [PSCustomObject]@{
                Path = "Get-AppxPackage -Name '$famName'"
                Hive = 'AppX (not ARP)'
                Note = 'MSIX/APPX packages do NOT register under HKLM\Software\...\Uninstall. Use the AppX PowerShell module; family-name path is the detection target.'
            }
        }

        default { return $null }
    }
}

# ---------------------------------------------------------------------------
# MSI Analysis
# ---------------------------------------------------------------------------

function Test-MsiModuleAvailable {
    <#
    .SYNOPSIS
        Checks if the MSI module is available -- currently loaded in the session
        (vendored under Lib\MSI\ and imported by the shell at startup) OR
        installed on the system (PSGallery).
    #>
    if ($null -eq $script:MsiModuleAvailable) {
        $loaded = $null -ne (Get-Module -Name MSI -ErrorAction SilentlyContinue)
        $onDisk = $null -ne (Get-Module -ListAvailable -Name MSI -ErrorAction SilentlyContinue)
        $script:MsiModuleAvailable = $loaded -or $onDisk
    }
    return $script:MsiModuleAvailable
}

function Get-MsiProperties {
    <#
    .SYNOPSIS
        Reads all properties from an MSI Property table.
    .DESCRIPTION
        Uses PSGallery MSI module if available, otherwise falls back to COM interop.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns an ordered hashtable of every MSI Property table row by design; Get-MsiProperty would imply a single named property.')]
    param([Parameter(Mandatory)][string]$MsiPath)

    Write-Log "Reading MSI properties: $MsiPath"

    # Try PSGallery MSI module first
    if (Test-MsiModuleAvailable) {
        try {
            Write-Log "Using MSI module for property extraction"
            Import-Module MSI -ErrorAction Stop
            $props = Get-MSIProperty -Path $MsiPath -ErrorAction Stop
            $result = [ordered]@{}
            foreach ($p in $props) {
                $result[$p.Property] = $p.Value
            }
            Write-Log "Read $($result.Count) MSI properties via MSI module"
            return $result
        }
        catch {
            Write-Log "MSI module failed, falling back to COM: $_" -Level WARN
        }
    }

    # COM interop fallback
    Write-Log "Using COM interop for MSI property extraction"
    $installer = $null; $db = $null; $view = $null; $record = $null

    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @($MsiPath, 0))

        $sql = 'SELECT Property, Value FROM Property'
        $view = $db.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $db, @($sql))
        $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null

        $result = [ordered]@{}
        while ($true) {
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
            if ($null -eq $record) { break }
            $propName = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
            $propValue = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 2)
            $result[$propName] = $propValue
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($record) | Out-Null
        }

        Write-Log "Read $($result.Count) MSI properties via COM"
        return $result
    }
    catch {
        Write-Log "Failed to read MSI properties: $_" -Level ERROR
        return [ordered]@{}
    }
    finally {
        foreach ($o in @($record, $view, $db, $installer)) {
            if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) {
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) | Out-Null
            }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Get-MsiSummaryInfo {
    <#
    .SYNOPSIS
        Reads MSI summary information stream (architecture, package code, etc.).
    #>
    param([Parameter(Mandatory)][string]$MsiPath)

    Write-Log "Reading MSI summary info: $MsiPath"

    # Try PSGallery MSI module first
    if (Test-MsiModuleAvailable) {
        try {
            Import-Module MSI -ErrorAction Stop
            $summary = Get-MSISummaryInfo -Path $MsiPath -ErrorAction Stop
            return [PSCustomObject]@{
                Template     = $summary.Template
                RevisionNumber = $summary.RevisionNumber
                Subject      = $summary.Subject
                Author       = $summary.Author
                Keywords     = $summary.Keywords
                Comments     = $summary.Comments
                Architecture = if ($summary.Template -match 'x64|Intel64|64') { 'x64' } elseif ($summary.Template -match 'Intel|x86') { 'x86' } else { $summary.Template }
            }
        }
        catch {
            Write-Log "MSI module summary failed, falling back to COM: $_" -Level WARN
        }
    }

    # COM fallback
    $installer = $null; $db = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @($MsiPath, 0))
        $summaryInfo = $db.GetType().InvokeMember("SummaryInformation", "GetProperty", $null, $db, @(0))

        $template = $summaryInfo.GetType().InvokeMember("Property", "GetProperty", $null, $summaryInfo, @(7))
        $revision = $summaryInfo.GetType().InvokeMember("Property", "GetProperty", $null, $summaryInfo, @(9))
        $subject  = $summaryInfo.GetType().InvokeMember("Property", "GetProperty", $null, $summaryInfo, @(3))
        $author   = $summaryInfo.GetType().InvokeMember("Property", "GetProperty", $null, $summaryInfo, @(4))

        return [PSCustomObject]@{
            Template       = [string]$template
            RevisionNumber = [string]$revision
            Subject        = [string]$subject
            Author         = [string]$author
            Keywords       = ''
            Comments       = ''
            Architecture   = if ([string]$template -match 'x64|Intel64|64') { 'x64' } elseif ([string]$template -match 'Intel|x86') { 'x86' } else { [string]$template }
        }
    }
    catch {
        Write-Log "Failed to read MSI summary: $_" -Level ERROR
        return $null
    }
    finally {
        foreach ($o in @($summaryInfo, $db, $installer)) {
            if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) {
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) | Out-Null
            }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Get-MspMetadata {
    <#
    .SYNOPSIS
        Reads the metadata stored in a Windows Installer Patch (.msp) file.
    .DESCRIPTION
        Opens the patch via OpenDatabase mode 32 (MSIOPEN_PATCHFILE) and
        returns a PSCustomObject with SummaryInformation (Template,
        RevisionNumber), MsiPatchMetadata (DisplayName, TargetProductName,
        Classification, ...), and an InferredDisplayVersion parsed from
        MsiPatchMetadata.DisplayName.

        Returns $null when the file is not a valid MSP or COM access fails.
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log "MSP file not found: $Path" -Level WARN
        return $null
    }

    Write-Log "Reading MSP metadata: $Path"

    $installerType = [Type]::GetTypeFromProgID('WindowsInstaller.Installer')
    if (-not $installerType) {
        Write-Log "WindowsInstaller.Installer COM ProgID not available" -Level ERROR
        return $null
    }

    $installer = $null
    $summary   = $null
    $db        = $null
    $view      = $null

    try {
        $installer = [Activator]::CreateInstance($installerType)

        # SummaryInformation reads via the installer (no DB open needed for
        # this stream). Template + RevisionNumber tell us the target
        # ProductCodes the patch applies to and the patch's own PatchCode(s).
        $summary = $installer.GetType().InvokeMember('SummaryInformation','GetProperty',$null,$installer,@($Path, 0))

        $template = ''
        $revision = ''
        $title    = ''
        $subject  = ''
        $author   = ''
        $keywords = ''
        $comments = ''
        try { $title    = [string]$summary.GetType().InvokeMember('Property','GetProperty',$null,$summary,@(2))  } catch { $null = $_ }
        try { $subject  = [string]$summary.GetType().InvokeMember('Property','GetProperty',$null,$summary,@(3))  } catch { $null = $_ }
        try { $author   = [string]$summary.GetType().InvokeMember('Property','GetProperty',$null,$summary,@(4))  } catch { $null = $_ }
        try { $keywords = [string]$summary.GetType().InvokeMember('Property','GetProperty',$null,$summary,@(5))  } catch { $null = $_ }
        try { $comments = [string]$summary.GetType().InvokeMember('Property','GetProperty',$null,$summary,@(6))  } catch { $null = $_ }
        try { $template = [string]$summary.GetType().InvokeMember('Property','GetProperty',$null,$summary,@(7))  } catch { $null = $_ }
        try { $revision = [string]$summary.GetType().InvokeMember('Property','GetProperty',$null,$summary,@(9))  } catch { $null = $_ }

        # Template is a semicolon-separated list of target ProductCodes.
        # Adobe and others ship one MSP that targets many language-variant
        # ProductCodes (same template, different lang code in the GUID).
        $targetProducts = @()
        if ($template) {
            foreach ($t in $template.Split(';')) {
                $tt = $t.Trim()
                if ($tt -match '^\{[0-9A-Fa-f-]{36}\}$') { $targetProducts += $tt }
            }
        }

        # RevisionNumber is a concatenated string of curly-brace GUIDs:
        # one PatchCode per target ProductCode listed in Template. They're
        # not separated by ';' -- they're back-to-back '{...}{...}{...}'.
        $patchCodes = @()
        if ($revision) {
            $regex = [regex]'\{[0-9A-Fa-f-]{36}\}'
            foreach ($m in $regex.Matches($revision)) { $patchCodes += $m.Value }
        }
        $primaryPatchCode = if ($patchCodes.Count -gt 0) { $patchCodes[0] } else { $null }

        # OpenDatabase mode 32 = MSIOPEN_PATCHFILE. Required to query patch tables.
        $db = $installer.GetType().InvokeMember('OpenDatabase','InvokeMethod',$null,$installer,@($Path, 32))

        # MsiPatchMetadata: StringData(1)=Company, (2)=Property, (3)=Value.
        # Every void-return InvokeMember below is wrapped in [void] so its
        # return value doesn't leak into this function's pipeline output.
        $meta = @{}
        try {
            $view = $db.GetType().InvokeMember('OpenView','InvokeMethod',$null,$db,@('SELECT * FROM `MsiPatchMetadata`'))
            [void]$view.GetType().InvokeMember('Execute','InvokeMethod',$null,$view,$null)
            while ($true) {
                $rec = $view.GetType().InvokeMember('Fetch','InvokeMethod',$null,$view,$null)
                if ($null -eq $rec) { break }
                try {
                    $prop = [string]$rec.GetType().InvokeMember('StringData','GetProperty',$null,$rec,@(2))
                    $val  = [string]$rec.GetType().InvokeMember('StringData','GetProperty',$null,$rec,@(3))
                    if ($prop) { $meta[$prop] = $val }
                } finally {
                    # Each Fetch returns a fresh COM record; releasing it
                    # per-iteration matches Get-MsiProperties' pattern and
                    # prevents the long-running shell from leaking RCWs
                    # across hundreds of analyses.
                    if ([System.Runtime.InteropServices.Marshal]::IsComObject($rec)) {
                        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($rec) } catch { $null = $_ }
                    }
                }
            }
            [void]$view.GetType().InvokeMember('Close','InvokeMethod',$null,$view,$null)
        } catch {
            Write-Log ("MsiPatchMetadata table read failed: " + $_.Exception.Message) -Level WARN
        }

        # Infer the post-patch DisplayVersion from MsiPatchMetadata.DisplayName.
        # The Windows Installer convention (followed by Adobe, Microsoft Office,
        # most enterprise vendors) is: DisplayName = "ProductName (version)".
        # Fall back to Description if DisplayName lacks a parenthetical.
        $patchDisplayName = [string]$meta['DisplayName']
        $patchDescription = [string]$meta['Description']
        $inferredVersion  = $null
        $verRegex = [regex]'\(([0-9]+(?:\.[0-9]+){1,3})\)'
        foreach ($src in @($patchDisplayName, $patchDescription, $subject, $title)) {
            if ($src) {
                $m = $verRegex.Match([string]$src)
                if ($m.Success) { $inferredVersion = $m.Groups[1].Value; break }
            }
        }

        return [PSCustomObject]@{
            PrimaryPatchCode       = $primaryPatchCode
            PatchCodes             = $patchCodes
            TargetProductCodes     = $targetProducts
            PatchDisplayName       = $patchDisplayName
            PatchDescription       = $patchDescription
            TargetProductName      = [string]$meta['TargetProductName']
            ManufacturerName       = [string]$meta['ManufacturerName']
            Classification         = [string]$meta['Classification']
            MoreInfoURL            = [string]$meta['MoreInfoURL']
            CreationTimeUTC        = [string]$meta['CreationTimeUTC']
            AllowRemoval           = [string]$meta['AllowRemoval']
            InferredDisplayVersion = $inferredVersion
            SummaryTitle           = $title
            SummarySubject         = $subject
            SummaryAuthor          = $author
            SummaryKeywords        = $keywords
            SummaryComments        = $comments
            Template               = $template
            RevisionNumber         = $revision
            MsiPatchMetadata       = $meta
        }
    }
    catch {
        Write-Log ("Failed to read MSP metadata: " + $_.Exception.Message) -Level WARN
        return $null
    }
    finally {
        foreach ($obj in @($view, $db, $summary, $installer)) {
            if ($null -ne $obj -and [System.Runtime.InteropServices.Marshal]::IsComObject($obj)) {
                try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null } catch { $null = $_ }
            }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# ---------------------------------------------------------------------------
# Deployment Field Resolution
# ---------------------------------------------------------------------------

function Get-DeploymentFields {
    <#
    .SYNOPSIS
        Resolves common Add/Remove Programs registry fields from available installer metadata.
    .DESCRIPTION
        Derives DisplayName, DisplayVersion, SilentUninstallString, Vendor, and the predicted
        UninstallRegistryKey from the best available source. Priority (highest first):
        1. PackageMetadata (from Get-ChocolateyMetadata / Get-IntunewinMetadata / Get-MsixManifest / Get-PsadtMetadata / Get-SquirrelMetadata / Get-WixBurnMetadata)
        2. MSI properties
        3. FileVersionInfo
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns a PSCustomObject aggregating multiple ARP fields; singular would imply a single-field lookup.')]
    param(
        [Parameter(Mandatory)][PSCustomObject]$FileInfo,
        [hashtable]$MsiProperties,
        [PSCustomObject]$Switches,
        [PSCustomObject]$PackageMetadata,
        [string]$InstallerType,
        # Optional. MSI Template architecture from Get-MsiSummaryInfo. Required for
        # accurate WOW6432Node routing on x86 MSIs because FileInfo.Architecture is
        # 'N/A (see MSI Summary)' for .msi files.
        [PSCustomObject]$MsiSummary
    )

    # DisplayName: PackageMetadata.DisplayName > MSI ProductName > FileVersionInfo ProductName > FileDescription
    $displayName = if ($PackageMetadata -and $PackageMetadata.DisplayName) {
        $PackageMetadata.DisplayName
    } elseif ($MsiProperties -and $MsiProperties.Contains('ProductName') -and $MsiProperties['ProductName']) {
        $MsiProperties['ProductName']
    } elseif ($FileInfo.ProductName) {
        $FileInfo.ProductName
    } elseif ($FileInfo.FileDescription) {
        $FileInfo.FileDescription
    } else { '' }

    # DisplayVersion: PackageMetadata.DisplayVersion > MSI ProductVersion > FileVersionInfo ProductVersion > FileVersion
    $displayVersion = if ($PackageMetadata -and $PackageMetadata.DisplayVersion) {
        $PackageMetadata.DisplayVersion
    } elseif ($MsiProperties -and $MsiProperties.Contains('ProductVersion') -and $MsiProperties['ProductVersion']) {
        $MsiProperties['ProductVersion']
    } elseif ($FileInfo.ProductVersion) {
        $FileInfo.ProductVersion
    } elseif ($FileInfo.FileVersion) {
        $FileInfo.FileVersion
    } else { '' }

    # SilentUninstallString: PackageMetadata.SilentUninstallCommand > Switches.Uninstall
    $silentUninstallString = if ($PackageMetadata -and $PackageMetadata.SilentUninstallCommand) {
        $PackageMetadata.SilentUninstallCommand
    } elseif ($Switches) {
        $Switches.Uninstall
    } else { '' }

    # Vendor: PackageMetadata.Publisher > MSI Manufacturer > FileVersionInfo CompanyName
    $vendor = if ($PackageMetadata -and $PackageMetadata.Publisher) {
        $PackageMetadata.Publisher
    } elseif ($MsiProperties -and $MsiProperties.Contains('Manufacturer') -and $MsiProperties['Manufacturer']) {
        $MsiProperties['Manufacturer']
    } elseif ($FileInfo.CompanyName) {
        $FileInfo.CompanyName
    } else { '' }

    # Trim every resolved field. Inno Setup (and some MSI tools) emit fixed-width
    # space-padded strings in the PE resource: Git for Windows ships DisplayName
    # as "Git" + ~60 spaces. Without trimming, the spaces leak into the predicted
    # UninstallRegistryKey ("HKLM:\...\Uninstall\Git                   _is1") and
    # break any downstream automation that uses the path verbatim.
    $displayName           = ([string]$displayName).Trim()
    $displayVersion        = ([string]$displayVersion).Trim()
    $vendor                = ([string]$vendor).Trim()
    $silentUninstallString = ([string]$silentUninstallString).Trim()

    $partial = [PSCustomObject]@{
        DisplayName          = $displayName
        DisplayVersion       = $displayVersion
        SilentUninstallString = $silentUninstallString
        Vendor               = $vendor
    }

    $uninstallKey = ''
    $uninstallKeyNote = ''
    if ($InstallerType) {
        $msiArch = if ($MsiSummary -and $MsiSummary.PSObject.Properties['Architecture']) { [string]$MsiSummary.Architecture } else { '' }
        $reg = Get-UninstallRegistryKey -InstallerType $InstallerType `
            -FileInfo $FileInfo -MsiProperties $MsiProperties `
            -PackageMetadata $PackageMetadata -DeploymentFields $partial `
            -MsiArchitecture $msiArch
        if ($reg) {
            $uninstallKey = [string]$reg.Path
            $uninstallKeyNote = [string]$reg.Note
        }
    }

    return [PSCustomObject]@{
        DisplayName              = $displayName
        DisplayVersion           = $displayVersion
        SilentUninstallString    = $silentUninstallString
        Vendor                   = $vendor
        UninstallRegistryKey     = $uninstallKey
        UninstallRegistryKeyNote = $uninstallKeyNote
    }
}

# ---------------------------------------------------------------------------
# Silent Switch Analysis
# ---------------------------------------------------------------------------

function Get-SilentSwitchDatabase {
    <#
    .SYNOPSIS
        Returns the static lookup table of known silent install switches per installer type.
    #>
    return @{
        'MSI' = @{
            Install   = 'msiexec.exe /i "<MSI>" /qn /norestart'
            Uninstall = 'msiexec.exe /x "<ProductCode>" /qn /norestart'
            Notes     = 'ProductCode required for uninstall. Add /L*v "<log>" for verbose logging.'
        }
        'NSIS' = @{
            Install   = '"<EXE>" /S'
            Uninstall = '"<UninstallEXE>" /S'
            Notes     = '/S is CASE SENSITIVE (uppercase S). Some NSIS installers also support /D=<path> for install directory.'
        }
        'InnoSetup' = @{
            Install   = '"<EXE>" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-'
            Uninstall = '"<UninstallEXE>" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART'
            Notes     = '/SP- suppresses the initial "This will install..." prompt. /DIR="<path>" for custom directory.'
        }
        'InstallShield' = @{
            Install   = '"<EXE>" /s /v"/qn"'
            Uninstall = '"<EXE>" /s /v"/qn" /x'
            Notes     = 'Some InstallShield installers require a response file: -s -f1"setup.iss"'
        }
        'WixBurn' = @{
            Install   = '"<EXE>" /quiet /norestart'
            Uninstall = '"<EXE>" /uninstall /quiet /norestart'
            Notes     = '/log "<path>" for logging. Burn bundles contain embedded MSI/EXE payloads.'
        }
        'AdvancedInstaller' = @{
            Install   = '"<EXE>" /i /qn'
            Uninstall = '"<EXE>" /x /qn'
            Notes     = 'Based on MSI technology. May also support msiexec.exe switches on the embedded MSI.'
        }
        '7zSFX' = @{
            Install   = 'Self-extracting archive. Extract with 7z.exe, then run the embedded installer.'
            Uninstall = 'Depends on the extracted payload.'
            Notes     = 'Not a true installer. Extract contents to identify the actual installer inside.'
        }
        'WinRarSFX' = @{
            Install   = 'Self-extracting archive. Extract with 7z.exe or WinRAR, then run the embedded installer.'
            Uninstall = 'Depends on the extracted payload.'
            Notes     = 'Not a true installer. Extract contents to identify the actual installer inside.'
        }
        'BitRock' = @{
            Install   = '"<EXE>" --mode unattended --unattendedmodeui none'
            Uninstall = '"<UninstallEXE>" --mode unattended --unattendedmodeui none'
            Notes     = 'BitRock InstallBuilder (used by PostgreSQL, Bitnami, JFrog Artifactory, etc.). Modes: unattended, text, gtk, win32, osx, xwindow. Run with --help for the full switch list on a specific product. Uninstaller lives at %INSTALLDIR%\uninstall.exe by default.'
        }
        'Chocolatey' = @{
            Install   = 'choco install <PackageId> --version=<Version> -y --source="<SourceDirOrFeed>"'
            Uninstall = 'choco uninstall <PackageId> -y'
            Notes     = 'Requires Chocolatey runtime (choco.exe). Package metadata lives in <PackageId>.nuspec at archive root; install script at tools/chocolateyInstall.ps1.'
        }
        'Intunewin' = @{
            Install   = '"<ExtractedSetup>" <OriginalSilentSwitches>'
            Uninstall = 'Defined in the Intune portal, not in the .intunewin file.'
            Notes     = 'Not directly executable. The inner IntunePackage.intunewin is AES-encrypted; decryption keys are delivered by Intune to the managed endpoint. Metadata lives in IntuneWinPackage/Metadata/Detection.xml; MsiInfo sub-element appears only when the source installer was an MSI.'
        }
        'Msix' = @{
            Install   = 'Add-AppxPackage -Path "<msix>"'
            Uninstall = 'Remove-AppxPackage -Package "<PackageFullName>"'
            Notes     = 'Per-user: Add-AppxPackage. All-users / provisioned: Add-AppxProvisionedPackage -Online -PackagePath "<msix>" -SkipLicense. PackageFullName = Name_Version_Arch_ResourceId_PublisherHash; retrieve via Get-AppxPackage after install. Applies to .msix and .appx.'
        }
        'MsixBundle' = @{
            Install   = 'Add-AppxPackage -Path "<msixbundle>"'
            Uninstall = 'Remove-AppxPackage -Package "<PackageFullName>"'
            Notes     = 'Same cmdlets as single MSIX. Bundle targets multiple architectures (x86/x64/ARM64) and/or language resource packs; AppxBundleManifest.xml lists which.'
        }
        'PsadtV3' = @{
            Install   = 'Deploy-Application.exe -DeploymentType Install -DeployMode Silent'
            Uninstall = 'Deploy-Application.exe -DeploymentType Uninstall -DeployMode Silent'
            Notes     = 'PowerShell App Deployment Toolkit v3. Wrapper script -- the Deploy-Application.ps1 file contains per-app install/uninstall logic. Pure .ps1 form: powershell.exe -ExecutionPolicy Bypass -File Deploy-Application.ps1 -DeploymentType Install -DeployMode Silent'
        }
        'PsadtV4' = @{
            Install   = 'Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent'
            Uninstall = 'Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent'
            Notes     = 'PowerShell App Deployment Toolkit v4. Toolkit logic lives in the PSAppDeployToolkit module; per-app logic in Invoke-AppDeployToolkit.ps1. Pure .ps1 form: powershell.exe -ExecutionPolicy Bypass -File Invoke-AppDeployToolkit.ps1 -DeploymentType Install -DeployMode Silent'
        }
        'Squirrel' = @{
            Install   = '"<Setup.exe>" --silent'
            Uninstall = '"%LOCALAPPDATA%\<AppName>\Update.exe" --uninstall -s'
            Notes     = 'Squirrel.Windows / Electron Setup.exe. Installs per-user to %LOCALAPPDATA%\<AppName>\. Use --silent (double-dash, lowercase), NOT the /S NSIS switch. Update.exe ships alongside the app and handles uninstalls and in-place updates.'
        }
        'NuGet' = @{
            Install   = 'nuget install <PackageId> -Version <Version> -Source "<SourceDirOrFeed>"'
            Uninstall = 'N/A (NuGet is a package source, not an installer)'
            Notes     = 'Plain NuGet package without Chocolatey tools. Not end-user installable by itself.'
        }
        'Unknown' = @{
            Install   = 'Unable to determine. Try: "<EXE>" /? or "<EXE>" --help for usage.'
            Uninstall = 'Check Add/Remove Programs for uninstall string.'
            Notes     = 'Installer framework not recognized. Common switches to try: /S, /silent, /quiet, /q, /VERYSILENT'
        }
    }
}

function Get-SilentSwitches {
    <#
    .SYNOPSIS
        Returns install/uninstall switches for the detected installer type.
    .DESCRIPTION
        Substitutes actual filename and ProductCode (if MSI) into the template strings.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns both install and uninstall switches on one object; plural matches the pair being returned.')]
    param(
        [Parameter(Mandatory)][string]$InstallerType,
        [Parameter(Mandatory)][string]$FilePath,
        [hashtable]$MsiProperties
    )

    $db = Get-SilentSwitchDatabase
    $fileName = Split-Path -Leaf $FilePath

    if (-not $db.Contains($InstallerType)) { $InstallerType = 'Unknown' }
    $entry = $db[$InstallerType]

    $install   = $entry.Install -replace '<EXE>', $fileName -replace '<MSI>', $fileName
    $uninstall = $entry.Uninstall -replace '<EXE>', $fileName -replace '<UninstallEXE>', 'uninstall.exe'

    if ($MsiProperties -and $MsiProperties.Contains('ProductCode')) {
        $install   = $install -replace '<ProductCode>', $MsiProperties['ProductCode']
        $uninstall = $uninstall -replace '<ProductCode>', $MsiProperties['ProductCode']
    }

    return [PSCustomObject]@{
        InstallerType = $InstallerType
        Install       = $install
        Uninstall     = $uninstall
        Notes         = $entry.Notes
    }
}

# ---------------------------------------------------------------------------
# Payload Extraction
# ---------------------------------------------------------------------------

function Find-7ZipPath {
    <#
    .SYNOPSIS
        Locates 7z.exe on the system.
    .DESCRIPTION
        Preferred path wins if valid; otherwise the standard Program Files
        locations; otherwise whatever `Get-Command 7z.exe` resolves from
        PATH (covers scoop / chocolatey / user-local installs).
    #>
    param([string]$PreferredPath)

    if ($PreferredPath -and (Test-Path -LiteralPath $PreferredPath)) { return $PreferredPath }

    $candidates = @(
        'C:\Program Files\7-Zip\7z.exe'
        'C:\Program Files (x86)\7-Zip\7z.exe'
    )
    # Also consider whatever is on PATH (covers non-standard install locations
    # like scoop, chocolatey user-local, portable copies).
    $fromPath = Get-Command '7z.exe' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1
    if ($fromPath) { $candidates += $fromPath }

    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return $c }
    }

    $inPath = Get-Command 7z.exe -ErrorAction SilentlyContinue
    if ($inPath) { return $inPath.Source }

    return $null
}

function Get-PayloadContents {
    <#
    .SYNOPSIS
        Lists contents of an installer using 7z.exe without extracting.
    .DESCRIPTION
        Returns array of PSCustomObjects with Name, Size, DateTime properties.
        Returns $null if 7z.exe is not found.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns the full contents listing from 7z l; plural matches the collection being returned.')]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$SevenZipPath
    )

    if (-not $SevenZipPath) { $SevenZipPath = Find-7ZipPath }
    if (-not $SevenZipPath) {
        Write-Log "7-Zip not found. Cannot list payload contents." -Level WARN
        return $null
    }

    Write-Log "Listing payload contents with 7z..."

    $tempOut = [System.IO.Path]::GetTempFileName()
    $tempErr = [System.IO.Path]::GetTempFileName()

    try {
        Start-Process -FilePath $SevenZipPath -ArgumentList @('l', "`"$Path`"") `
            -Wait -NoNewWindow `
            -RedirectStandardOutput $tempOut -RedirectStandardError $tempErr

        $output = Get-Content -LiteralPath $tempOut -ErrorAction SilentlyContinue

        # Parse 7z list output -- look for the file table between dashed lines
        $results = @()
        $inTable = $false
        $dashCount = 0

        foreach ($line in $output) {
            if ($line -match '^-{10,}') {
                $dashCount++
                if ($dashCount -eq 1) { $inTable = $true; continue }
                if ($dashCount -eq 2) { break }
            }
            if ($inTable -and $line.Trim()) {
                # 7-Zip list format: "[YYYY-MM-DD HH:MM:SS] Attr Size Compressed Name"
                # The date/time columns are EMPTY (just whitespace) for entries
                # without a known timestamp -- e.g. the gzip-compressed Inno
                # Setup wizard stream "[0]~" 7z surfaces for Git/Inno installers.
                # First try with date+time, fall back to no-date format.
                $attr = ''; $size = -1L; $name = ''
                if ($line -match '^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\s+(\S+)\s+(\d+)\s+\d*\s+(.+)$') {
                    $attr = $Matches[1]
                    $size = [long]$Matches[2]
                    $name = $Matches[3].Trim()
                } elseif ($line -match '^\s+(\.+|D\.+|[\.A-Z]+)\s+(\d+)\s+\d*\s+(.+)$') {
                    $attr = $Matches[1]
                    $size = [long]$Matches[2]
                    $name = $Matches[3].Trim()
                }
                if ($name) {
                    $results += [PSCustomObject]@{
                        Name = $name
                        Size = $size
                        SizeFormatted = if ($size -ge 1MB) { "{0:N1} MB" -f ($size / 1MB) } elseif ($size -ge 1KB) { "{0:N0} KB" -f ($size / 1KB) } else { "$size B" }
                        IsDirectory = ($attr -match 'D')
                    }
                }
            }
        }

        Write-Log "Found $($results.Count) items in payload"
        return $results
    }
    catch {
        Write-Log "Failed to list payload: $_" -Level ERROR
        return @()
    }
    finally {
        Remove-Item -LiteralPath $tempOut, $tempErr -ErrorAction SilentlyContinue
    }
}

function Get-EmbeddedInstallers {
    <#
    .SYNOPSIS
        Classifies installer-class entries in a payload listing.
    .DESCRIPTION
        Filters the rows from Get-PayloadContents down to entries that can
        themselves be analyzed (inner MSI / MSP / sub-EXE / CAB / .nupkg).
        Per-Type filters drop MSI internal streams and other noise.

        Returns an array of PSCustomObjects (Name, Kind, Size, SizeFormatted,
        Description, IsAnalyzable, ParentPath). Empty array is a valid result.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns a collection.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification='PackageMetadata is part of the public signature; reserved for richer per-type classification.')]
    param(
        [object[]]$Payload,
        [string]$Type = 'Unknown',
        $PackageMetadata = $null,
        [string]$ParentPath = ''
    )

    $null = $PackageMetadata  # Reserved for Phase 2 (WiX Burn manifest).

    if (-not $Payload) { return ,@() }

    # Per-Type prefilter: MSI files surface their internal table streams via 7-Zip
    # (!Property, !File, Binary.X, Icon.Y, [5]SummaryInformation, numbered streams
    # with control-char prefixes). None are installable; drop them at the source.
    $candidates = @($Payload)
    if ($Type -eq 'MSI') {
        $candidates = @($candidates | Where-Object {
            $n = [string]$_.Name
            (-not $_.IsDirectory) -and
            (-not [string]::IsNullOrWhiteSpace($n)) -and
            (-not ($n.StartsWith('!'))) -and
            (-not ($n.StartsWith('['))) -and
            (-not ($n.StartsWith('Binary.'))) -and
            (-not ($n.StartsWith('Icon.'))) -and
            (-not ($n -match '^[\x01-\x1F]'))
        })
    }

    $kindOrder = @{
        'MSI'   = 0
        'MSP'   = 1
        'MSIX'  = 2
        'APPX'  = 3
        'EXE'   = 4
        'NUPKG' = 5
        'ZIP'   = 6
        '7Z'    = 7
        'CAB'   = 8
    }

    $results = @()
    foreach ($entry in $candidates) {
        if ($entry.IsDirectory) { continue }

        $name = [string]$entry.Name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        # 7-Zip surfaces some synthetic stream names with angle-bracket
        # decoration (e.g. <inner.msi> when the MSI was embedded without a
        # real filename, [0]~ for Inno wizard). Strip the brackets just for
        # extension classification; keep the original Name on the row for
        # downstream 7z e -i!"<name>" filtering compatibility.
        $extName = $name.Trim('<','>')
        $ext  = [System.IO.Path]::GetExtension($extName).ToLowerInvariant()
        $kind = $null
        $desc = ''

        switch ($ext) {
            '.msi'         { $kind = 'MSI';   $desc = 'Windows Installer database' }
            '.msp'         { $kind = 'MSP';   $desc = 'Windows Installer patch' }
            '.cab'         { $kind = 'CAB';   $desc = 'Cabinet archive' }
            '.exe'         { $kind = 'EXE';   $desc = 'Executable (re-detected on drill-down)' }
            '.zip'         { $kind = 'ZIP';   $desc = 'Zip archive' }
            '.nupkg'       { $kind = 'NUPKG'; $desc = 'NuGet / Chocolatey package' }
            '.appx'        { $kind = 'APPX';  $desc = 'AppX package' }
            '.appxbundle'  { $kind = 'APPX';  $desc = 'AppX bundle' }
            '.msix'        { $kind = 'MSIX';  $desc = 'MSIX package' }
            '.msixbundle'  { $kind = 'MSIX';  $desc = 'MSIX bundle' }
            '.7z'          { $kind = '7Z';    $desc = '7-Zip archive' }
        }

        if (-not $kind) { continue }

        # Per-Type descriptions: the same .cab means very different things inside
        # an MSI (Media-table feature CAB) vs. inside a WiX Burn bundle (chained
        # package container) vs. inside a custom PE wrapper (component payload).
        switch ($Type) {
            'MSI' {
                if ($kind -eq 'CAB') { $desc = 'MSI Media-table feature CAB' }
            }
            'WixBurn' {
                if ($kind -in @('MSI','EXE','MSP')) { $desc = 'Burn bundle chained package' }
                elseif ($kind -eq 'CAB') { $desc = 'Burn container CAB' }
            }
            'Squirrel' {
                if ($kind -eq 'NUPKG') { $desc = 'Squirrel app payload (.nupkg)' }
                elseif ($name -ieq 'Update.exe') { $desc = 'Squirrel updater stub' }
            }
            'NSIS' {
                if ($kind -eq 'EXE' -and $name -match '^uninst|^Un_?A|uninst\.exe$') { $desc = 'NSIS uninstaller stub' }
                if ($kind -eq 'CAB') { $desc = 'NSIS-bundled CAB payload' }
            }
            'PsadtV3' { if ($kind -in @('MSI','EXE','MSP')) { $desc = 'PSADT Files\ payload' } }
            'PsadtV4' { if ($kind -in @('MSI','EXE','MSP')) { $desc = 'PSADT Files\ payload' } }
            'Chocolatey' { if ($kind -in @('MSI','EXE')) { $desc = 'Chocolatey tools\ payload' } }
        }

        $sizeFormatted = if ($entry.PSObject.Properties['SizeFormatted']) {
            [string]$entry.SizeFormatted
        } else {
            $s = [long]$entry.Size
            if ($s -ge 1MB) { '{0:N1} MB' -f ($s / 1MB) }
            elseif ($s -ge 1KB) { '{0:N0} KB' -f ($s / 1KB) }
            else { "$s B" }
        }

        $results += [PSCustomObject]@{
            Name          = $name
            Kind          = $kind
            Size          = [long]$entry.Size
            SizeFormatted = $sizeFormatted
            Description   = $desc
            IsAnalyzable  = $true
            ParentPath    = $ParentPath
        }
    }

    # Bring wrapped / chained installers to the top, then component CABs and
    # archives by Kind, then size descending within each kind. Mirrors the
    # mental model: "tell me what this thing IS before showing me its parts."
    $sorted = $results | Sort-Object @{ Expression = {
        if ($kindOrder.ContainsKey($_.Kind)) { $kindOrder[$_.Kind] } else { 99 }
    } }, @{ Expression = { -[long]$_.Size } }

    return ,@($sorted)
}

function Expand-PayloadEntry {
    <#
    .SYNOPSIS
        Extracts ONE named entry from an archive into a flat output dir.
    .DESCRIPTION
        Returns the full path of the extracted file; $null on failure.
        The include argument is quoted at the argv level so entry names
        containing spaces are passed to 7-Zip as a single token.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Wraps 7z.exe; OutputDir is caller-owned temp.')]
    param(
        [Parameter(Mandatory)][string]$SevenZipPath,
        [Parameter(Mandatory)][string]$ArchivePath,
        [Parameter(Mandatory)][string]$EntryName,
        [Parameter(Mandatory)][string]$OutputDir
    )

    if (-not (Test-Path -LiteralPath $SevenZipPath)) {
        Write-Log "7-Zip not found at: $SevenZipPath" -Level WARN
        return $null
    }
    if (-not (Test-Path -LiteralPath $ArchivePath)) {
        Write-Log "Archive not found: $ArchivePath" -Level WARN
        return $null
    }
    if (-not (Test-Path -LiteralPath $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    # CRITICAL: -i!Name MUST be quoted at the argv level. Without the quote
    # 7-Zip splits the argument on any whitespace in EntryName, silently
    # mis-matching the include pattern and emitting "No files to process"
    # with exit code 0. See module comments above for the bug history.
    $arglist = @(
        'e',
        "`"$ArchivePath`"",
        "-o`"$OutputDir`"",
        "`"-i!$EntryName`"",
        '-y',
        '-r'
    )
    $proc = Start-Process -FilePath $SevenZipPath -ArgumentList $arglist -Wait -NoNewWindow -PassThru
    if ($null -eq $proc -or $proc.ExitCode -ne 0) {
        $code = if ($proc) { $proc.ExitCode } else { 'no proc' }
        Write-Log ("7z exit code {0} extracting '{1}' from '{2}'" -f $code, $EntryName, $ArchivePath) -Level WARN
        return $null
    }

    # Strip the angle-bracket synthetic decoration (e.g. <inner.msi>) the
    # same way Get-EmbeddedInstallers does, and pull the leaf out of any
    # nested path.
    $leaf = [System.IO.Path]::GetFileName($EntryName.Trim('<','>'))
    $candidate = Join-Path $OutputDir $leaf
    if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }

    # 7z sometimes sanitizes path components; fall back to "newest file
    # in OutputDir" as a best-effort recovery.
    $newest = Get-ChildItem -LiteralPath $OutputDir -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($newest) { return $newest.FullName }
    return $null
}

function Expand-InstallerPayload {
    <#
    .SYNOPSIS
        Extracts installer contents to a directory using 7z.exe.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$OutputPath,
        [string]$SevenZipPath
    )

    if (-not $SevenZipPath) { $SevenZipPath = Find-7ZipPath }
    if (-not $SevenZipPath) {
        Write-Log "7-Zip not found. Cannot extract payload." -Level ERROR
        return $null
    }

    if (-not $OutputPath) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        $OutputPath = Join-Path $env:TEMP "InstallerAnalysis\$baseName`_$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    }

    if (-not (Test-Path -LiteralPath $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    Write-Log "Extracting to: $OutputPath"

    $proc = Start-Process -FilePath $SevenZipPath -ArgumentList @('x', "`"$Path`"", "-o`"$OutputPath`"", '-y') `
        -Wait -NoNewWindow -PassThru

    if ($proc.ExitCode -eq 0) {
        Write-Log "Extraction complete"
    } else {
        Write-Log "7z.exe exited with code $($proc.ExitCode)" -Level WARN
    }

    return $OutputPath
}

# ---------------------------------------------------------------------------
# String Analysis
# ---------------------------------------------------------------------------

function Get-BinaryStrings {
    <#
    .SYNOPSIS
        Extracts printable strings from a binary file, both ASCII and UTF-16LE.
    .DESCRIPTION
        Modern Windows-native installers (Squirrel/Electron, recent MSI custom
        actions, NSIS Unicode builds, Inno Setup Unicode, .NET-based bootstrappers)
        keep most of their interesting strings -- paths, URLs, registry keys --
        as UTF-16LE wide chars. An ASCII-only scan misses all of them. This
        function collects BOTH encodings from the read window.

        Default 16MB window (was 2MB). The interesting bootstrapper region of
        large installers (GitHub Desktop is 181MB) lives at the start; raising
        from 2MB to 16MB catches it without blowing memory on multi-GB files.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns the collection of printable strings found in the binary.')]
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$MinLength = 8,
        [int]$MaxBytes = 16MB
    )

    $fileSize = (Get-Item -LiteralPath $Path).Length
    $readSize = [Math]::Min($MaxBytes, $fileSize)

    $stream = $null
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        $bytes = New-Object byte[] $readSize
        $stream.Read($bytes, 0, $readSize) | Out-Null
    }
    finally {
        if ($stream) { try { $stream.Close() } catch { $null = $_ } }
    }

    # Per-byte PowerShell loops are ~minutes-slow at 16MB. Decode the whole
    # buffer once per encoding (.NET native, sub-second) then let a compiled
    # regex find printable runs of >= MinLength. Same result, ~1000x faster.

    # ASCII pass. .NET ASCII.GetString maps any byte > 0x7F to '?', which the
    # printable-range regex correctly rejects. Bytes 0x00-0x1F decode to control
    # chars and also get rejected. Net result: runs of printable ASCII only.
    $asciiText = [System.Text.Encoding]::ASCII.GetString($bytes)
    $printable = [regex]::new("[\x20-\x7E]{$MinLength,}",
        [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $strings = New-Object 'System.Collections.Generic.List[string]'
    foreach ($m in $printable.Matches($asciiText)) { [void]$strings.Add($m.Value) }

    # UTF-16LE pass for wide-char strings (Squirrel paths, NSIS Unicode, modern
    # MSI custom action strings, etc.). Unicode.GetString interprets every byte
    # pair as a UTF-16LE code unit; the same printable-run regex pulls out the
    # readable substrings. Try both byte alignments since interesting wide
    # regions inside a binary aren't guaranteed to start on an even offset.
    $wideAligned   = [System.Text.Encoding]::Unicode.GetString($bytes, 0, ($bytes.Length - ($bytes.Length % 2)))
    foreach ($m in $printable.Matches($wideAligned)) { [void]$strings.Add($m.Value) }
    if ($bytes.Length -gt 2) {
        $wideUnaligned = [System.Text.Encoding]::Unicode.GetString($bytes, 1, ($bytes.Length - 1) - (($bytes.Length - 1) % 2))
        foreach ($m in $printable.Matches($wideUnaligned)) { [void]$strings.Add($m.Value) }
    }

    return $strings.ToArray()
}

function Get-InterestingStrings {
    <#
    .SYNOPSIS
        Filters binary strings into categorized interesting findings.
    .DESCRIPTION
        Scans the binary for interesting patterns (URLs, registry keys, file
        paths, GUIDs, framework markers, version triples) across both ASCII and
        UTF-16LE encodings. Operates directly on decoded buffers via compiled
        regex -- no intermediate "all strings" array. Modern installers
        (Squirrel/Electron, MSI custom actions, Unicode NSIS/Inno) store these
        as wide chars, so the ASCII-only scan that lived here before missed
        nearly everything on Windows-native bootstrappers.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns a categorized collection of interesting strings.')]
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$MaxBytes = 16MB
    )

    Write-Log "Scanning for interesting strings..."

    $empty = @{ InstallerMarkers = @(); URLs = @(); RegistryPaths = @(); FilePaths = @(); GUIDs = @(); VersionStrings = @() }
    if (-not (Test-Path -LiteralPath $Path)) { return $empty }

    $fileSize = (Get-Item -LiteralPath $Path).Length
    $readSize = [Math]::Min($MaxBytes, $fileSize)
    $stream = $null
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        $bytes = New-Object byte[] $readSize
        $stream.Read($bytes, 0, $readSize) | Out-Null
    }
    catch { return $empty }
    finally {
        if ($stream) { try { $stream.Close() } catch { $null = $_ } }
    }

    # Decode in two views: ASCII (one byte per char) and UTF-16LE at even
    # alignment. The odd-alignment pass was dropped after profiling -- tripled
    # the regex work, produced zero hits on every binary in the test set.
    $asciiView = [System.Text.Encoding]::ASCII.GetString($bytes)
    $wideEven  = [System.Text.Encoding]::Unicode.GetString($bytes, 0, ($bytes.Length - ($bytes.Length % 2)))
    $views = @($asciiView, $wideEven)

    # ONE combined regex per view. Iterating PowerShell `foreach ($m in ...)`
    # over MatchCollection wraps every Match in a PSObject -- that wrap was the
    # bottleneck (1.5-2.3s per pattern per view, 12 passes total = ~20s on a
    # 13MB MSI). The Match.NextMatch() chain stays in pure .NET. Combined with
    # alternation + named groups, one pass categorises everything.
    $opts = [System.Text.RegularExpressions.RegexOptions]::Compiled -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    # URL char class is bounded to actual valid URL characters per RFC 3986
    # (alnum + ":/?#@!$&'()*+,;=._~%-"). Critically excludes backslash and curly
    # braces so the URL branch can't accidentally swallow an adjacent registry
    # path or GUID when the combined alternation runs.
    # Registry/file char classes likewise tightened to stop at characters that
    # almost certainly mark the end of the value.
    $combined = [regex]::new(
        '(?<url>https?://[A-Za-z0-9:/?#@!$&''()*+,;=._~%-]+)' +
        '|(?<reg>(?:HKLM|HKCU|HKCR|HKU|HKEY_LOCAL_MACHINE|HKEY_CURRENT_USER|HKEY_CLASSES_ROOT|HKEY_USERS)\\[A-Za-z0-9._\\ -]+)' +
        '|(?<file>(?:[A-Za-z]:\\|%[A-Za-z_]+%\\|\$\{[A-Za-z_]+\}\\)[A-Za-z0-9._\\ -]+\.(?:exe|msi|msp|dll|sys|cab|nupkg|ps1|js|json|xml|cfg|ini|config|cmd|bat|reg|zip|7z|wxs|wxl))' +
        '|(?<guid>\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\})' +
        '|(?<marker>NullsoftInst|Inno Setup|InstallShield|WixBurn|Advanced Installer|WixBundleManifest|Microsoft Visual C\+\+|\.NET Framework|NSIS|SquirrelTemp|squirrel-(?:install|updated|uninstall|firstrun|obsolete)|BitRock)' +
        '|(?<ver>\b\d+\.\d+\.\d+(?:\.\d+)?\b)',
        $opts)

    $markersSet  = New-Object 'System.Collections.Generic.HashSet[string]'
    $urlsSet     = New-Object 'System.Collections.Generic.HashSet[string]'
    $regPathsSet = New-Object 'System.Collections.Generic.HashSet[string]'
    $filesSet    = New-Object 'System.Collections.Generic.HashSet[string]'
    $guidsSet    = New-Object 'System.Collections.Generic.HashSet[string]'
    $versionsSet = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($view in $views) {
        if ([string]::IsNullOrEmpty($view)) { continue }
        $m = $combined.Match($view)
        while ($m.Success) {
            if     ($m.Groups['url'].Success)    { [void]$urlsSet.Add($m.Value) }
            elseif ($m.Groups['reg'].Success)    { [void]$regPathsSet.Add($m.Value) }
            elseif ($m.Groups['file'].Success)   { [void]$filesSet.Add($m.Value) }
            elseif ($m.Groups['guid'].Success)   { [void]$guidsSet.Add($m.Value.ToUpperInvariant()) }
            elseif ($m.Groups['marker'].Success) { [void]$markersSet.Add($m.Value) }
            elseif ($m.Groups['ver'].Success)    { [void]$versionsSet.Add($m.Value) }
            $m = $m.NextMatch()
        }
    }

    $markers   = @($markersSet  | Sort-Object)
    $urls      = @($urlsSet     | Sort-Object)
    $regPaths  = @($regPathsSet | Sort-Object)
    $filePaths = @($filesSet    | Sort-Object)
    $guids     = @($guidsSet    | Sort-Object)
    $versions  = @($versionsSet | Sort-Object | Select-Object -First 50)

    Write-Log "Found: $($markers.Count) markers, $($urls.Count) URLs, $($regPaths.Count) registry paths, $($filePaths.Count) file paths, $($guids.Count) GUIDs"

    return @{
        InstallerMarkers = $markers
        URLs             = $urls
        RegistryPaths    = $regPaths
        FilePaths        = $filePaths
        GUIDs            = $guids
        VersionStrings   = $versions
    }
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

function Export-AnalysisReport {
    <#
    .SYNOPSIS
        Exports analysis results as CSV (property/value pairs).
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $rows = @()
    foreach ($row in $DataTable.Rows) {
        $obj = [ordered]@{}
        foreach ($col in $DataTable.Columns) { $obj[$col.ColumnName] = $row[$col.ColumnName] }
        $rows += [PSCustomObject]$obj
    }
    $rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Exported CSV to $OutputPath"
}

function Export-AnalysisHtml {
    <#
    .SYNOPSIS
        Exports analysis results as a styled HTML report.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$ReportTitle = 'Installer Analysis Report'
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $css = @(
        '<style>',
        'body { font-family: "Segoe UI", Arial, sans-serif; margin: 20px; background: #fafafa; }',
        'h1 { color: #0078D4; margin-bottom: 4px; }',
        '.summary { color: #666; margin-bottom: 12px; font-size: 0.9em; }',
        'table { border-collapse: collapse; width: 100%; margin-top: 12px; }',
        'th { background: #0078D4; color: #fff; padding: 8px 12px; text-align: left; }',
        'td { padding: 6px 12px; border-bottom: 1px solid #e0e0e0; }',
        'tr:nth-child(even) { background: #f5f5f5; }',
        '.prop { font-weight: bold; width: 200px; }',
        '</style>'
    ) -join "`r`n"

    $headerRow = ($DataTable.Columns | ForEach-Object { "<th>$($_.ColumnName)</th>" }) -join ''
    $bodyRows = foreach ($row in $DataTable.Rows) {
        $cells = foreach ($col in $DataTable.Columns) {
            $val = [string]$row[$col.ColumnName]
            $cls = if ($col.Ordinal -eq 0) { ' class="prop"' } else { '' }
            "<td$cls>$val</td>"
        }
        "<tr>$($cells -join '')</tr>"
    }

    $html = @(
        '<!DOCTYPE html>', '<html><head><meta charset="utf-8"><title>' + $ReportTitle + '</title>',
        $css, '</head><body>', "<h1>$ReportTitle</h1>",
        "<div class='summary'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Rows: $($DataTable.Rows.Count)</div>",
        "<table><thead><tr>$headerRow</tr></thead>",
        "<tbody>$($bodyRows -join "`r`n")</tbody></table>",
        '</body></html>'
    ) -join "`r`n"

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    Write-Log "Exported HTML to $OutputPath"
}

function script:Expand-InstallerPlaceholders {
    <#
    .SYNOPSIS
        Substitutes known concrete values into the templated strings that the
        silent-switch database, package-metadata extractors, and uninstall-key
        notes carry (e.g. "<AppName>", "<Setup.exe>", "<ProductCode>").
    .DESCRIPTION
        The tool is for inspecting installers, so every value the analysis
        already knows should appear baked into the rendered output -- not as a
        generic placeholder. This helper takes a final rendered string and a
        context hashtable of name->value substitutions and replaces every
        "<name>" token that appears verbatim.

        Substitutions only happen for keys with non-empty values. Empty values
        leave the placeholder intact so the reader still sees it's missing.
    #>
    param(
        [string]$Text,
        [System.Collections.IDictionary]$Context
    )
    if ([string]::IsNullOrEmpty($Text) -or -not $Context) { return $Text }
    foreach ($key in $Context.Keys) {
        $val = [string]$Context[$key]
        if (-not [string]::IsNullOrEmpty($val)) {
            $Text = $Text.Replace("<$key>", $val)
        }
    }
    return $Text
}

function New-AnalysisSummaryText {
    <#
    .SYNOPSIS
        Returns a clipboard-ready text summary of the analysis, including any
        format-specific PACKAGE METADATA section for modern package types.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Pure function that builds a string; does not touch external state.')]
    param(
        [Parameter(Mandatory)][PSCustomObject]$FileInfo,
        [Parameter(Mandatory)][string]$InstallerType,
        [PSCustomObject]$Switches,
        [hashtable]$MsiProperties,
        [PSCustomObject]$DeploymentFields,
        [PSCustomObject]$PackageMetadata,
        [object[]]$MspMetadata,
        [PSCustomObject]$InnerMsiData
    )

    # File header: facts about the source binary that don't appear elsewhere.
    # Version/Company are intentionally NOT here -- Deployment Fields has
    # DisplayVersion/Vendor which carry the same values (the canonical app-level
    # facts, possibly overridden by PackageMetadata for Squirrel etc).
    $lines = @(
        "Installer Analysis: $($FileInfo.FileName)",
        ("-" * 50),
        "Type:         $InstallerType",
        "Architecture: $($FileInfo.Architecture)",
        "Size:         $($FileInfo.FileSizeFormatted)",
        "SHA-256:      $($FileInfo.SHA256)",
        "Signed:       $($FileInfo.SignatureStatus)$(if ($FileInfo.SignerSubject) { " ($($FileInfo.SignerSubject))" })"
    )

    # SilentInstall is computed the same way ConvertTo-DeploymentJson does it:
    # PackageMetadata.SilentInstallCommand > Switches.Install. Same for uninstall.
    # This used to live at the top as "Silent Install: / Silent Uninstall:" lines
    # AND again under Deployment Fields as "SilentUninstallString:"; folded into
    # one Deployment Fields block to stop the duplicate.
    $silentInstallStr = ''
    if ($Switches -and $Switches.PSObject.Properties['Install']) { $silentInstallStr = [string]$Switches.Install }
    if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['SilentInstallCommand'] -and $PackageMetadata.SilentInstallCommand) {
        $silentInstallStr = [string]$PackageMetadata.SilentInstallCommand
    }

    if ($DeploymentFields) {
        $lines += ""
        $lines += "Deployment Fields:"
        $lines += "  DisplayName:           $($DeploymentFields.DisplayName)"
        $lines += "  DisplayVersion:        $($DeploymentFields.DisplayVersion)"
        $lines += "  Vendor:                $($DeploymentFields.Vendor)"
        if ($silentInstallStr)            { $lines += "  SilentInstallString:   $silentInstallStr" }
        $lines += "  SilentUninstallString: $($DeploymentFields.SilentUninstallString)"
        if ($DeploymentFields.PSObject.Properties['UninstallRegistryKey'] -and $DeploymentFields.UninstallRegistryKey) {
            $lines += "  UninstallRegistryKey:  $($DeploymentFields.UninstallRegistryKey)"
        }
    }

    # Effective Post-Patch Detection Target. Surfaced when the outer file
    # contains at least one .msp (Adobe Reader, Office security rollups,
    # vendor base-MSI + cumulative-MSP shipments). The MSP's
    # MsiPatchMetadata.DisplayName carries the post-patch DisplayVersion the
    # ARP key will hold once the patch applies; the inner MSI's ProductCode
    # is the ARP key name. Together they give a MECM-ready detection rule
    # WITHOUT installing the product.
    $primaryMsp = $null
    if ($MspMetadata) {
        $mspArr = @($MspMetadata)
        if ($mspArr.Count -gt 0) { $primaryMsp = $mspArr[0] }
    }
    if ($primaryMsp) {
        $lines += ""
        $lines += "Effective Post-Patch Detection Target:"
        if ($primaryMsp.TargetProductName) {
            $lines += "  Patched product:         $($primaryMsp.TargetProductName)"
        }
        $detectionPc = ''
        if ($InnerMsiData -and $InnerMsiData.ProductCode) {
            $detectionPc = [string]$InnerMsiData.ProductCode
        } elseif ($primaryMsp.TargetProductCodes -and @($primaryMsp.TargetProductCodes).Count -gt 0) {
            $detectionPc = [string]@($primaryMsp.TargetProductCodes)[0]
        }
        if ($detectionPc) {
            $lines += "  ProductCode:             $detectionPc"
        }
        if ($primaryMsp.InferredDisplayVersion) {
            $lines += "  DisplayVersion:          $($primaryMsp.InferredDisplayVersion)"
        }
        if ($InnerMsiData -and $InnerMsiData.ProductVersion -and $InnerMsiData.ProductVersion -ne $primaryMsp.InferredDisplayVersion) {
            $lines += "  BaseDisplayVersion:      $($InnerMsiData.ProductVersion)"
        }
        if ($primaryMsp.Classification) {
            $lines += "  Patch Classification:    $($primaryMsp.Classification)"
        }
        if ($detectionPc) {
            $arch = ''
            if ($InnerMsiData -and $InnerMsiData.PSObject.Properties['MsiArchitecture']) {
                $arch = [string]$InnerMsiData.MsiArchitecture
            }
            $is32 = ($arch -match '^(x86|Intel)$')
            $is64 = ($arch -match '^(x64|Intel64|AMD64)$')
            if ($is32) {
                $lines += "  UninstallRegistryKey:    HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$detectionPc"
            } elseif ($is64) {
                $lines += "  UninstallRegistryKey:    HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$detectionPc"
            } else {
                $lines += "  UninstallRegistryKey x86: HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$detectionPc"
                $lines += "  UninstallRegistryKey x64: HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$detectionPc"
            }
            if ($primaryMsp.InferredDisplayVersion) {
                $lines += "  ARP value match:         DisplayVersion = $($primaryMsp.InferredDisplayVersion)"
            }
        }
        $mspArrCount = @($MspMetadata).Count
        if ($mspArrCount -gt 1) {
            $lines += "  Patches detected:        $mspArrCount"
        }
        if ($primaryMsp.PSObject.Properties['SourceFile'] -and $primaryMsp.SourceFile) {
            $lines += "  Source MSP:              $($primaryMsp.SourceFile)"
        }
        if ($primaryMsp.MoreInfoURL) {
            $lines += "  More info URL:           $($primaryMsp.MoreInfoURL)"
        }
    }

    if ($MsiProperties -and $MsiProperties.Count -gt 0) {
        $lines += ""
        $lines += "MSI Properties:"
        if ($MsiProperties.Contains('ProductCode'))  { $lines += "  Product Code:  $($MsiProperties['ProductCode'])" }
        if ($MsiProperties.Contains('UpgradeCode'))  { $lines += "  Upgrade Code:  $($MsiProperties['UpgradeCode'])" }
        if ($MsiProperties.Contains('ProductVersion')) { $lines += "  Version:       $($MsiProperties['ProductVersion'])" }
        if ($MsiProperties.Contains('Manufacturer'))  { $lines += "  Manufacturer:  $($MsiProperties['Manufacturer'])" }
    }

    if ($PackageMetadata) {
        $pkg = $PackageMetadata
        $pkgType = if ($pkg.PSObject.Properties['InstallerType']) { [string]$pkg.InstallerType } else { $InstallerType }
        $hasAnyPkg = $false

        # Per-type Package Metadata blocks: surface only fields that are NOT
        # already in Deployment Fields. DisplayName / DisplayVersion / Vendor
        # equivalents are dropped from these blocks since the canonical values
        # live in the Deployment Fields section directly above.
        switch ($pkgType) {
            { $_ -in 'Chocolatey','NuGet' } {
                $lines += ""; $lines += "Package Metadata (nuspec):"
                if ($pkg.PSObject.Properties['PackageId'])  { $lines += "  Id:            $($pkg.PackageId)" }
                if ($pkg.PSObject.Properties['ProjectUrl']) { $lines += "  Project URL:   $($pkg.ProjectUrl)" }
                if ($pkg.PSObject.Properties['Tags'])        { $lines += "  Tags:          $($pkg.Tags)" }
                if ($pkg.PSObject.Properties['Description'] -and $pkg.Description) {
                    $desc = [string]$pkg.Description
                    if ($desc.Length -gt 140) { $desc = $desc.Substring(0,137) + '...' }
                    $lines += "  Description:   $desc"
                }
                $hasAnyPkg = $true
            }
            'Intunewin' {
                $lines += ""; $lines += "Package Metadata (IntuneWinPackage\Metadata\Detection.xml):"
                if ($pkg.PSObject.Properties['SetupFile'])    { $lines += "  Setup File:    $($pkg.SetupFile)" }
                if ($pkg.PSObject.Properties['ToolVersion'])  { $lines += "  Tool Version:  $($pkg.ToolVersion)" }
                if ($pkg.PSObject.Properties['EncryptionInfo'] -and $pkg.EncryptionInfo) {
                    $lines += "  Encrypted:     yes (payload keys live in Intune; cannot decrypt locally)"
                }
                if ($pkg.PSObject.Properties['MsiInfo'] -and $pkg.MsiInfo) {
                    $mi = $pkg.MsiInfo
                    $lines += "  Source MSI:"
                    foreach ($prop in 'MsiProductCode','MsiProductVersion','MsiUpgradeCode','MsiExecutionContext','MsiRequiresReboot') {
                        if ($mi.PSObject.Properties[$prop] -and $mi.$prop) {
                            $lines += ("    {0,-22} {1}" -f ($prop + ':'), $mi.$prop)
                        }
                    }
                }
                $hasAnyPkg = $true
            }
            { $_ -in 'Msix','MsixBundle' } {
                $lines += ""; $lines += "Package Metadata (AppxManifest):"
                # Identity.Publisher is the cryptographic publisher (CN=...) --
                # different from Deployment Fields Vendor (friendly name), so
                # keep it. Name/Version dup DisplayName/DisplayVersion -- drop.
                if ($pkg.PSObject.Properties['Identity'] -and $pkg.Identity) {
                    $id = $pkg.Identity
                    $lines += "  Identity:"
                    foreach ($key in 'Publisher','ProcessorArchitecture','ResourceId') {
                        if ($id.$key) { $lines += ("    {0,-22} {1}" -f ($key + ':'), $id.$key) }
                    }
                }
                if ($pkg.PSObject.Properties['PropertiesDescription'] -and $pkg.PropertiesDescription) {
                    $lines += "  Description:   $($pkg.PropertiesDescription)"
                }
                if ($pkgType -eq 'MsixBundle' -and $pkg.PSObject.Properties['BundledPackages']) {
                    $bp = @($pkg.BundledPackages)
                    $lines += "  Bundled Packages: $($bp.Count)"
                    foreach ($b in $bp) {
                        $tok = @()
                        if ($b.Type)         { $tok += "[$($b.Type)]" }
                        if ($b.Architecture) { $tok += "arch=$($b.Architecture)" }
                        if ($b.Version)      { $tok += "v$($b.Version)" }
                        if ($b.ResourceId)   { $tok += "resource=$($b.ResourceId)" }
                        $tok += ($b.FileName)
                        $lines += ("    - " + ($tok -join '  '))
                    }
                }
                $hasAnyPkg = $true
            }
            { $_ -in 'PsadtV3','PsadtV4' } {
                $lines += ""; $lines += "Package Metadata (PSAppDeployToolkit $($pkg.ToolkitVariant) header):"
                if ($pkg.PSObject.Properties['ToolkitVersion'] -and $pkg.ToolkitVersion) {
                    $lines += "  Toolkit Ver:   $($pkg.ToolkitVersion)"
                }
                # AppVendor/AppName/AppVersion dup Vendor/DisplayName/DisplayVersion.
                # Keep the PSADT-specific fields: AppArch, AppLang, AppRevision,
                # plus script provenance (ScriptVersion/Date/Author).
                if ($pkg.PSObject.Properties['AppMetadata'] -and $pkg.AppMetadata) {
                    $am = $pkg.AppMetadata
                    $lines += "  App Header:"
                    foreach ($field in 'AppArch','AppLang','AppRevision','ScriptVersion','ScriptDate','ScriptAuthor') {
                        # AppMetadata is an ordered dictionary from ConvertFrom-PsadtDeployApplication
                        # ([ordered]@{}) which is System.Collections.IDictionary but NOT [hashtable].
                        # PSCustomObject AppMetadata also possible; accept both shapes.
                        $val = if ($am -is [System.Collections.IDictionary]) { $am[$field] } elseif ($am.PSObject.Properties[$field]) { $am.$field } else { '' }
                        if ($val) { $lines += ("    {0,-18} {1}" -f ($field + ':'), $val) }
                    }
                }
                $hasAnyPkg = $true
            }
            'MSP' {
                $lines += ""; $lines += "Patch Metadata (MsiPatchMetadata):"
                if ($pkg.PSObject.Properties['PatchDisplayName'] -and $pkg.PatchDisplayName) {
                    $lines += "  DisplayName:             $($pkg.PatchDisplayName)"
                }
                if ($pkg.PSObject.Properties['TargetProductName'] -and $pkg.TargetProductName) {
                    $lines += "  Target Product:          $($pkg.TargetProductName)"
                }
                if ($pkg.PSObject.Properties['InferredDisplayVersion'] -and $pkg.InferredDisplayVersion) {
                    $lines += "  Post-patch Version:      $($pkg.InferredDisplayVersion)  (parsed from DisplayName)"
                }
                if ($pkg.PSObject.Properties['Classification'] -and $pkg.Classification) {
                    $lines += "  Classification:          $($pkg.Classification)"
                }
                if ($pkg.PSObject.Properties['ManufacturerName'] -and $pkg.ManufacturerName) {
                    $lines += "  Manufacturer:            $($pkg.ManufacturerName)"
                }
                if ($pkg.PSObject.Properties['MoreInfoURL'] -and $pkg.MoreInfoURL) {
                    $lines += "  More info URL:           $($pkg.MoreInfoURL)"
                }
                if ($pkg.PSObject.Properties['CreationTimeUTC'] -and $pkg.CreationTimeUTC) {
                    $lines += "  Created (UTC):           $($pkg.CreationTimeUTC)"
                }
                if ($pkg.PSObject.Properties['PrimaryPatchCode'] -and $pkg.PrimaryPatchCode) {
                    $lines += "  Primary PatchCode:       $($pkg.PrimaryPatchCode)"
                }
                if ($pkg.PSObject.Properties['PatchCodes'] -and $pkg.PatchCodes) {
                    $pcCount = @($pkg.PatchCodes).Count
                    if ($pcCount -gt 1) { $lines += "  Patch Codes total:       $pcCount (one per target product variant)" }
                }
                if ($pkg.PSObject.Properties['TargetProductCodes'] -and $pkg.TargetProductCodes) {
                    $tpcs = @($pkg.TargetProductCodes)
                    $lines += "  Target ProductCodes:     $($tpcs.Count) variant(s)"
                    $maxShow = [Math]::Min(5, $tpcs.Count)
                    for ($i = 0; $i -lt $maxShow; $i++) {
                        $lines += "    [$i] $($tpcs[$i])"
                    }
                    if ($tpcs.Count -gt $maxShow) {
                        $lines += "    ... ($($tpcs.Count - $maxShow) more; full list in JSON export)"
                    }
                }
                $hasAnyPkg = $true
            }
            'Squirrel' {
                $lines += ""; $lines += "Package Metadata (Squirrel / Electron Setup.exe):"
                # DisplayName / DisplayVersion dropped here -- both appear under
                # Deployment Fields directly above. What's left is Squirrel-specific:
                # the markers we matched, the embedded nupkg refs, and the detection
                # confidence.
                if ($pkg.PSObject.Properties['MarkersFound'] -and $pkg.MarkersFound) {
                    $markers = @($pkg.MarkersFound)
                    $lines += "  Markers:       $($markers.Count) found -> $($markers -join ', ')"
                }
                # Render the actual NupkgReferences array Get-SquirrelMetadata
                # produces. The previous "EmbeddedNupkg" check was for a field
                # name that never existed in the metadata object, so nupkg refs
                # were silently dropped from the Overview tab.
                if ($pkg.PSObject.Properties['NupkgReferences'] -and $pkg.NupkgReferences) {
                    $refs = @($pkg.NupkgReferences)
                    if ($refs.Count -gt 0) {
                        $primary = $refs | Where-Object { $_.Kind -eq 'full' } | Select-Object -First 1
                        if (-not $primary) { $primary = $refs[0] }
                        $lines += "  Embedded nupkg: $($primary.FileName) [$($primary.Kind)]"
                        if ($refs.Count -gt 1) {
                            $extras = ($refs | Where-Object { $_.FileName -ne $primary.FileName } |
                                      Select-Object -First 5 |
                                      ForEach-Object { "$($_.FileName) [$($_.Kind)]" })
                            $lines += "  Other refs:    $($refs.Count - 1) more -> $($extras -join ', ')"
                        }
                    }
                }
                if ($pkg.PSObject.Properties['Confidence'] -and $pkg.Confidence) {
                    $lines += "  Confidence:    $($pkg.Confidence)"
                }
                $hasAnyPkg = $true
            }
            default { }
        }

        if (-not $hasAnyPkg) {
            # Generic fallback: dump well-known cross-format fields if present.
            $lines += ""; $lines += "Package Metadata:"
            foreach ($prop in 'DisplayName','DisplayVersion','Publisher','Architecture','ProductCodeOrEquivalent') {
                if ($pkg.PSObject.Properties[$prop] -and $pkg.$prop) {
                    $lines += ("  {0,-14} {1}" -f ($prop + ':'), $pkg.$prop)
                }
            }
        }
    }

    # Bake every value we already know into the rendered output. This is an
    # installer-analysis tool -- placeholders like "<AppName>" or "<Setup.exe>"
    # in the displayed text are a lie when we have the concrete value sitting
    # right there in the analysis result.
    $ctx = [ordered]@{}
    if ($FileInfo) {
        if ($FileInfo.FileName) {
            $ctx['Setup.exe']    = [string]$FileInfo.FileName
            $ctx['MSI']          = [string]$FileInfo.FileName
            $ctx['EXE']          = [string]$FileInfo.FileName
            $ctx['UninstallEXE'] = [string]$FileInfo.FileName  # NSIS/Inno fallback if unknown
        }
    }
    # AppName: prefer the cleanest source -- PackageMetadata.ProductCodeOrEquivalent
    # (Squirrel/PSADT use the raw app id here), else DeploymentFields.DisplayName.
    $appNameForSubs = ''
    if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['ProductCodeOrEquivalent'] -and $PackageMetadata.ProductCodeOrEquivalent) {
        $appNameForSubs = [string]$PackageMetadata.ProductCodeOrEquivalent
    }
    if (-not $appNameForSubs -and $DeploymentFields -and $DeploymentFields.DisplayName) {
        $appNameForSubs = [string]$DeploymentFields.DisplayName
    }
    if ($appNameForSubs) {
        $ctx['AppName'] = $appNameForSubs
        $ctx['AppId']   = $appNameForSubs
    }
    if ($DeploymentFields -and $DeploymentFields.DisplayVersion) {
        $ctx['Version'] = [string]$DeploymentFields.DisplayVersion
    }
    if ($MsiProperties) {
        if ($MsiProperties.Contains('ProductCode') -and $MsiProperties['ProductCode']) {
            $ctx['ProductCode'] = [string]$MsiProperties['ProductCode']
        }
        if ($MsiProperties.Contains('UpgradeCode') -and $MsiProperties['UpgradeCode']) {
            $ctx['UpgradeCode'] = [string]$MsiProperties['UpgradeCode']
        }
    }
    # Squirrel-specific: the uninstall EXE is a known absolute path once AppName resolves.
    if ($appNameForSubs -and $InstallerType -eq 'Squirrel') {
        $ctx['UninstallEXE'] = "%LOCALAPPDATA%\$appNameForSubs\Update.exe"
    }

    $rendered = ($lines -join "`r`n")
    return Expand-InstallerPlaceholders -Text $rendered -Context $ctx
}

function ConvertTo-DeploymentJson {
    <#
    .SYNOPSIS
        Builds a MECM-packaging-friendly JSON digest of the analysis result.
    .DESCRIPTION
        Returns a single JSON string suitable for pasting into packaging scripts or
        the clipboard. Shape is a flat digest of the fields a packager needs:
        source-file metadata, application identity, deploy commands, and a detection
        hint keyed to the installer type. Format-specific extras are nested under
        Raw.PackageMetadata; MSI properties are nested under Raw.MsiProperties.

        Values default to empty strings (never $null) so downstream scripts can rely
        on the schema. Pair with Get-InstallerFileInfo / Get-DeploymentFields /
        Get-SilentSwitches / Get-MsiProperties / Get-*Metadata to fill the inputs.
    .EXAMPLE
        $json = ConvertTo-DeploymentJson -FileInfo $fi -InstallerType 'NSIS' `
            -Switches $sw -DeploymentFields $df
        Set-Clipboard -Value $json
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Pure function that builds a JSON string; does not touch external state.')]
    param(
        [Parameter(Mandatory)][PSCustomObject]$FileInfo,
        [Parameter(Mandatory)][string]$InstallerType,
        [PSCustomObject]$Switches,
        [hashtable]$MsiProperties,
        [PSCustomObject]$DeploymentFields,
        [PSCustomObject]$PackageMetadata,
        [int]$Depth = 6
    )

    $productCode = ''
    $upgradeCode = ''
    if ($MsiProperties) {
        if ($MsiProperties.Contains('ProductCode') -and $MsiProperties['ProductCode']) { $productCode = [string]$MsiProperties['ProductCode'] }
        if ($MsiProperties.Contains('UpgradeCode') -and $MsiProperties['UpgradeCode']) { $upgradeCode = [string]$MsiProperties['UpgradeCode'] }
    }
    if (-not $productCode -and $PackageMetadata -and $PackageMetadata.PSObject.Properties['ProductCodeOrEquivalent']) {
        $productCode = [string]$PackageMetadata.ProductCodeOrEquivalent
    }

    $displayName = if ($DeploymentFields -and $DeploymentFields.DisplayName) { [string]$DeploymentFields.DisplayName }
                   elseif ($PackageMetadata -and $PackageMetadata.PSObject.Properties['DisplayName']) { [string]$PackageMetadata.DisplayName }
                   elseif ($FileInfo.ProductName) { [string]$FileInfo.ProductName }
                   elseif ($FileInfo.FileDescription) { [string]$FileInfo.FileDescription }
                   else { '' }

    $displayVersion = if ($DeploymentFields -and $DeploymentFields.DisplayVersion) { [string]$DeploymentFields.DisplayVersion }
                      elseif ($PackageMetadata -and $PackageMetadata.PSObject.Properties['DisplayVersion']) { [string]$PackageMetadata.DisplayVersion }
                      elseif ($FileInfo.ProductVersion) { [string]$FileInfo.ProductVersion }
                      elseif ($FileInfo.FileVersion) { [string]$FileInfo.FileVersion }
                      else { '' }

    $publisher = if ($DeploymentFields -and $DeploymentFields.Vendor) { [string]$DeploymentFields.Vendor }
                 elseif ($PackageMetadata -and $PackageMetadata.PSObject.Properties['Publisher']) { [string]$PackageMetadata.Publisher }
                 elseif ($FileInfo.CompanyName) { [string]$FileInfo.CompanyName }
                 else { '' }

    $architecture = if ($FileInfo.Architecture) { [string]$FileInfo.Architecture }
                    elseif ($PackageMetadata -and $PackageMetadata.PSObject.Properties['Architecture']) { [string]$PackageMetadata.Architecture }
                    else { '' }

    $installCmd = ''
    $uninstallCmd = ''
    $notes = ''
    if ($Switches) {
        if ($Switches.PSObject.Properties['Install'])   { $installCmd   = [string]$Switches.Install }
        if ($Switches.PSObject.Properties['Uninstall']) { $uninstallCmd = [string]$Switches.Uninstall }
        if ($Switches.PSObject.Properties['Notes'])     { $notes        = [string]$Switches.Notes }
    }
    if ($DeploymentFields -and $DeploymentFields.SilentUninstallString) {
        $uninstallCmd = [string]$DeploymentFields.SilentUninstallString
    }
    if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['SilentInstallCommand'] -and $PackageMetadata.SilentInstallCommand) {
        $installCmd = [string]$PackageMetadata.SilentInstallCommand
    }
    if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['SilentUninstallCommand'] -and $PackageMetadata.SilentUninstallCommand) {
        $uninstallCmd = [string]$PackageMetadata.SilentUninstallCommand
    }

    $detectionHint = switch ($InstallerType) {
        'MSI'        { if ($productCode) { "MSI ProductCode detection: $productCode" } else { 'MSI detection: use ProductCode from the MSI Property table' } }
        'NSIS'       { 'Registry uninstall key detection (HKLM\...\Uninstall\<DisplayName>) or file-version on the primary EXE' }
        'InnoSetup'  { 'Registry uninstall key detection (HKLM\...\Uninstall\<AppId>_is1) or file-version on the primary EXE' }
        'InstallShield' { 'Registry uninstall key detection or file-version on the primary EXE' }
        'WixBurn'    { 'BundleUpgradeCode or related MSI ProductCode under HKLM\...\Uninstall' }
        'BitRock'    { 'Registry uninstall key detection keyed on DisplayName (BitRock writes HKLM\...\Uninstall\<AppName> by default)' }
        'Msix'       { if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['Identity']) { "MSIX family-name detection via Get-AppxPackage -Name $($PackageMetadata.Identity.Name)" } else { 'MSIX family-name detection via Get-AppxPackage' } }
        'MsixBundle' { if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['Identity']) { "MSIX bundle family-name detection via Get-AppxPackage -Name $($PackageMetadata.Identity.Name)" } else { 'MSIX bundle family-name detection via Get-AppxPackage' } }
        'Chocolatey' { if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['PackageId']) { "Chocolatey package detection: choco list --local-only $($PackageMetadata.PackageId)" } else { 'Chocolatey package detection via choco list --local-only' } }
        'NuGet'      { 'NuGet-package detection via nuget list or the consumer package manager' }
        'Intunewin'  { 'Intune-managed detection defined in the Intune portal; see Detection.xml for the source installer hints' }
        'PsadtV3'    { 'Registry uninstall key detection keyed on AppName/AppVersion (configure in Deploy-Application.ps1)' }
        'PsadtV4'    { 'Registry uninstall key detection keyed on AppName/AppVersion (configure in Invoke-AppDeployToolkit.ps1)' }
        'Squirrel'   { 'Registry uninstall key in HKCU\...\Uninstall\<AppName> (per-user) or file-version on the installed <AppName>.exe' }
        '7zSFX'      { 'Not a true installer; detection depends on the extracted payload' }
        'WinRarSFX'  { 'Not a true installer; detection depends on the extracted payload' }
        default      { 'Pick a file-version or registry-uninstall detection per the extracted payload' }
    }

    # Substitute every placeholder we can resolve before emitting JSON. Without
    # this, downstream consumers see "<Setup.exe>" / "<AppName>" / "<ProductCode>"
    # even though the analysis already knows the concrete values. Mirrors what
    # New-AnalysisSummaryText does for the Overview text -- the two outputs must
    # stay consistent.
    $jsonCtx = [ordered]@{}
    if ($FileInfo -and $FileInfo.FileName) {
        $jsonCtx['Setup.exe']    = [string]$FileInfo.FileName
        $jsonCtx['MSI']          = [string]$FileInfo.FileName
        $jsonCtx['EXE']          = [string]$FileInfo.FileName
        $jsonCtx['UninstallEXE'] = [string]$FileInfo.FileName
    }
    $jsonAppName = ''
    if ($PackageMetadata -and $PackageMetadata.PSObject.Properties['ProductCodeOrEquivalent'] -and $PackageMetadata.ProductCodeOrEquivalent) {
        $jsonAppName = [string]$PackageMetadata.ProductCodeOrEquivalent
    }
    if (-not $jsonAppName -and $displayName) { $jsonAppName = $displayName }
    if ($jsonAppName) { $jsonCtx['AppName'] = $jsonAppName; $jsonCtx['AppId'] = $jsonAppName }
    # DisplayName is the placeholder used by NSIS / BitRock / Inno detection hints.
    # Without this key, the hint text leaks "<DisplayName>" even when we resolved it.
    if ($displayName) { $jsonCtx['DisplayName'] = $displayName }
    if ($displayVersion) { $jsonCtx['Version'] = $displayVersion }
    if ($productCode) { $jsonCtx['ProductCode'] = $productCode }
    if ($upgradeCode) { $jsonCtx['UpgradeCode'] = $upgradeCode }
    if ($jsonAppName -and $InstallerType -eq 'Squirrel') {
        $jsonCtx['UninstallEXE'] = "%LOCALAPPDATA%\$jsonAppName\Update.exe"
    }

    $installCmd     = Expand-InstallerPlaceholders -Text $installCmd     -Context $jsonCtx
    $uninstallCmd   = Expand-InstallerPlaceholders -Text $uninstallCmd   -Context $jsonCtx
    $notes          = Expand-InstallerPlaceholders -Text $notes          -Context $jsonCtx
    $detectionHint  = Expand-InstallerPlaceholders -Text ([string]$detectionHint) -Context $jsonCtx
    $regKeyNote     = if ($DeploymentFields -and $DeploymentFields.PSObject.Properties['UninstallRegistryKeyNote']) {
        Expand-InstallerPlaceholders -Text ([string]$DeploymentFields.UninstallRegistryKeyNote) -Context $jsonCtx
    } else { '' }

    $blob = [ordered]@{
        SchemaVersion    = '1.0'
        Source           = [ordered]@{
            FileName = if ($FileInfo.FileName) { [string]$FileInfo.FileName } else { '' }
            FileSize = if ($FileInfo.FileSize) { [long]$FileInfo.FileSize } else { 0 }
            SHA256   = if ($FileInfo.SHA256)   { [string]$FileInfo.SHA256 } else { '' }
        }
        Application      = [ordered]@{
            DisplayName    = $displayName
            DisplayVersion = $displayVersion
            Publisher      = $publisher
            Architecture   = $architecture
            InstallerType  = $InstallerType
            ProductCode    = $productCode
            UpgradeCode    = $upgradeCode
        }
        Deployment       = [ordered]@{
            InstallCommand   = $installCmd
            UninstallCommand = $uninstallCmd
            Notes            = $notes
        }
        Detection        = [ordered]@{
            Hint                     = $detectionHint
            UninstallRegistryKey     = if ($DeploymentFields -and $DeploymentFields.PSObject.Properties['UninstallRegistryKey']) { [string]$DeploymentFields.UninstallRegistryKey } else { '' }
            UninstallRegistryKeyNote = $regKeyNote
        }
        Raw              = [ordered]@{
            MsiProperties   = if ($MsiProperties) { $MsiProperties } else { [ordered]@{} }
            # PackageMetadata is copied through into Raw with placeholder
            # substitution applied to the known free-text command fields
            # (SilentInstallCommand, SilentUninstallCommand). Originals like
            # MarkersFound, NupkgReferences, Identity, etc. are passed through
            # untouched -- only fields that explicitly carry "<AppName>" /
            # "<Setup.exe>" template strings get expanded.
            PackageMetadata = if ($PackageMetadata) {
                $rawPkg = [ordered]@{}
                foreach ($prop in $PackageMetadata.PSObject.Properties) {
                    $val = $prop.Value
                    if ($prop.Name -in @('SilentInstallCommand','SilentUninstallCommand') -and $val -is [string]) {
                        $val = Expand-InstallerPlaceholders -Text ([string]$val) -Context $jsonCtx
                    }
                    $rawPkg[$prop.Name] = $val
                }
                [PSCustomObject]$rawPkg
            } else { $null }
        }
    }

    return ($blob | ConvertTo-Json -Depth $Depth)
}
