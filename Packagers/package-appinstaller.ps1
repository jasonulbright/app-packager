<#
Vendor: Microsoft
App: App Installer (WinGet)
CMName: App Installer (WinGet)
VendorUrl: https://github.com/microsoft/winget-cli
CPE: cpe:2.3:a:microsoft:app_installer:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/microsoft/winget-cli/releases
DownloadPageUrl: https://github.com/microsoft/winget-cli/releases
SecurityAdvisoryUrl: https://github.com/microsoft/winget-cli/security/advisories
IconSource: External
WsusSupport: No (InstallerTypeUnsupported|CustomInstall)
ArpName: none

.SYNOPSIS
    Packages App Installer (WinGet) for all users for ConfigMgr.

.DESCRIPTION
    Downloads the newest stable release of the Windows Package Manager from
    the winget-cli GitHub releases: the App Installer msixbundle, the
    dependency archive and the license file. Each file is checked against the
    SHA-256 digest that GitHub publishes for the asset, and the bundle must
    carry a valid Microsoft signature.

    The install script provisions the bundle for all users with
    Add-AppxProvisionedPackage, the dependencies for the processor
    architecture and the license file included. It needs no Microsoft Store
    and no network access. A user gets the package at the next sign-in; the
    script does not register it for a user who is signed in.

    Detection reads the provisioned package version from the all-user store
    and accepts the packaged version or a newer one, so a device that the
    Store updated stays compliant.

    The uninstall script removes the package for all users and the
    provisioning. A profile that is not loaded cannot be changed, so its copy
    can remain; detection reads the provisioning only.

    Supports two-phase operation:
      -StageOnly    Download, verify, extract the dependencies, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\App Installer (WinGet)\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 10

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Queries the winget-cli releases, outputs the version of the newest stable
    release that has all the files, and exits. No download or ConfigMgr
    changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
    - RBAC permissions to create Applications and Deployment Types
    - Write access to FileServerPath
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [ValidateSet('Nested','Flat')]
    [string]$ContentLayout = "Nested",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 10,
    [int]$MaximumRuntimeMins = 30,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly,
    [switch]$VerboseLog
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force -ErrorAction Stop
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$GitHubApiUrl = "https://api.github.com/repos/microsoft/winget-cli/releases?per_page=30"

$VendorFolder = "Microsoft"
$AppFolder    = "App Installer (WinGet)"
$AppName      = "App Installer (WinGet)"
$Publisher    = "Microsoft Corporation"

$PackageIdentity = "Microsoft.DesktopAppInstaller"
$BundleFileName  = "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
$DepsZipName     = "DesktopAppInstaller_Dependencies.zip"
$DepsJsonName    = "DesktopAppInstaller_Dependencies.json"
$Architectures   = @('x64', 'arm64')

$BaseDownloadRoot = Join-Path $DownloadRoot "AppInstaller"

# --- Functions ---


function Select-WinGetRelease {
    <#
    .SYNOPSIS
        Picks the newest stable release that lists every file the install needs.

    .DESCRIPTION
        A draft, a prerelease and a tag with a suffix (v1.30.140-preview) never
        match. An asset without a SHA-256 digest is not used, because its
        download cannot be verified. The license file name starts with a GUID,
        so it is chosen by its suffix.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Json)

    $releases = @(ConvertFrom-Json $Json | ForEach-Object { $_ })
    foreach ($release in $releases) {
        if ($release.draft -or $release.prerelease) { continue }
        $tag = [string]$release.tag_name
        if ($tag -notmatch '^v(?<v>\d+(?:\.\d+){2,3})$') { continue }
        $version = $Matches['v']

        $find = {
            param([string]$Pattern)
            foreach ($asset in @($release.assets)) {
                if ([string]$asset.name -match $Pattern -and [string]$asset.digest -match '^sha256:(?<h>[0-9a-fA-F]{64})$') {
                    return [pscustomobject]@{
                        Name        = [string]$asset.name
                        DownloadUrl = [string]$asset.browser_download_url
                        Sha256      = $Matches['h'].ToUpperInvariant()
                    }
                }
            }
            return $null
        }

        $bundle  = & $find '^Microsoft\.DesktopAppInstaller_8wekyb3d8bbwe\.msixbundle$'
        $depsZip = & $find '^DesktopAppInstaller_Dependencies\.zip$'
        $depsJson = & $find '^DesktopAppInstaller_Dependencies\.json$'
        $license = & $find '^[0-9a-fA-F]+_License1\.xml$'
        if (-not ($bundle -and $depsZip -and $depsJson -and $license)) { continue }

        return [pscustomobject]@{
            Version  = $version
            Tag      = $tag
            Bundle   = $bundle
            DepsZip  = $depsZip
            DepsJson = $depsJson
            License  = $license
        }
    }
    throw "No stable winget-cli release lists the App Installer bundle, the dependency archive and the license file with SHA-256 digests."
}


function Get-WinGetRelease {
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    $json = (curl.exe -L --fail --silent --show-error @(Get-GitHubApiCurlArgs) $GitHubApiUrl) -join ''
    if ($LASTEXITCODE -ne 0) { throw "Failed to query the winget-cli releases API." }

    $release = Select-WinGetRelease -Json $json
    Write-Log "Latest stable release        : $($release.Tag)" -Quiet:$Quiet
    return $release
}


function Get-BundleIdentity {
    <#
    .SYNOPSIS
        Returns the Name and Version of the bundle from AppxBundleManifest.xml.
    #>
    param([Parameter(Mandatory)][string]$Path)

    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry('AppxMetadata/AppxBundleManifest.xml')
        if (-not $entry) { throw "AppxBundleManifest.xml not found in the bundle: $Path" }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        try { [xml]$manifest = $reader.ReadToEnd() } finally { $reader.Dispose() }
    }
    finally { $zip.Dispose() }

    return $manifest.Bundle.Identity
}


function Get-VerifiedAsset {
    <#
    .SYNOPSIS
        Downloads an asset to the cache folder and checks it against its digest.
    #>
    param(
        [Parameter(Mandatory)]$Asset,
        [Parameter(Mandatory)][string]$Folder
    )

    $local = Join-Path $Folder $Asset.Name
    if (Test-Path -LiteralPath $local) {
        $cached = (Get-FileHash -LiteralPath $local -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($cached -ne $Asset.Sha256) {
            Write-Log "Local $($Asset.Name) does not match the published digest. Downloading again." -Level WARN
            Remove-Item -LiteralPath $local -Force -ErrorAction Stop
        }
    }
    if (-not (Test-Path -LiteralPath $local)) {
        Write-Log "Downloading $($Asset.Name)..."
        Invoke-DownloadWithRetry -Url $Asset.DownloadUrl -OutFile $local
    }
    $actual = (Get-FileHash -LiteralPath $local -Algorithm SHA256 -ErrorAction Stop).Hash
    if ($actual -ne $Asset.Sha256) {
        Remove-Item -LiteralPath $local -Force -ErrorAction SilentlyContinue
        throw "SHA256 mismatch for $($Asset.Name) (release digest: $($Asset.Sha256), file: $actual)."
    }
    Write-Log "SHA256 matches the digest    : $($Asset.Name)"
    return $local
}


function Expand-Dependencies {
    <#
    .SYNOPSIS
        Extracts the per-architecture dependency packages and checks them
        against the dependency list of the release.
    #>
    param(
        [Parameter(Mandatory)][string]$ZipPath,
        [Parameter(Mandatory)][string]$JsonPath,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string[]]$Architecture
    )

    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

    $dependencies = @((Get-Content -LiteralPath $JsonPath -Raw | ConvertFrom-Json).Dependencies)
    if ($dependencies.Count -eq 0) { throw "The dependency list is empty: $JsonPath" }

    $root = [System.IO.Path]::GetFullPath($Destination).TrimEnd('\') + '\'
    $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        foreach ($arch in $Architecture) {
            $archFolder = Join-Path $Destination $arch
            Initialize-Folder -Path $archFolder
            foreach ($dep in $dependencies) {
                $wanted = '{0}_{1}_{2}.appx' -f $dep.Name, $dep.Version, $arch
                $entry = @($zip.Entries | Where-Object { $_.FullName -ieq "$arch/$wanted" }) | Select-Object -First 1
                if (-not $entry) { throw "The dependency archive has no $arch/$wanted." }
                $target = [System.IO.Path]::GetFullPath((Join-Path $archFolder $wanted))
                if (-not $target.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
                    throw "A dependency entry extracts outside the stage folder: $($entry.FullName)"
                }
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $target, $true)
                Write-Log "Dependency staged            : $arch\$wanted"
            }
        }
    }
    finally { $zip.Dispose() }
}


function New-AppInstallerInstallScript {
    param([Parameter(Mandatory)][string]$BundleFileName)

    return (@(
        '$ErrorActionPreference = ''Stop''',
        'try {',
        '    if ([Environment]::OSVersion.Version -lt [version]''10.0.17763'') { throw ''App Installer needs Windows 10 version 1809 (build 17763) or later.'' }',
        '    $processor = if ($env:PROCESSOR_ARCHITEW6432) { $env:PROCESSOR_ARCHITEW6432 } else { $env:PROCESSOR_ARCHITECTURE }',
        '    $folder = switch ($processor) { ''AMD64'' { ''x64'' } ''ARM64'' { ''arm64'' } default { throw "App Installer is not packaged for the $processor processor architecture." } }',
        ('    $bundle = Join-Path $PSScriptRoot ''{0}''' -f $BundleFileName),
        '    $dependencies = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot (''deps\'' + $folder)) -Filter ''*.appx'' | ForEach-Object { $_.FullName })',
        '    if ($dependencies.Count -eq 0) { throw "No dependency packages for $folder in the content folder." }',
        '    $license = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter ''*_License1.xml'' | ForEach-Object { $_.FullName }) | Select-Object -First 1',
        '    if (-not $license) { throw ''The license file is missing from the content folder.'' }',
        '    Add-AppxProvisionedPackage -Online -PackagePath $bundle -DependencyPackagePath $dependencies -LicensePath $license | Out-Null',
        '}',
        'catch {',
        '    [Console]::Error.WriteLine($_.Exception.Message)',
        '    exit 1',
        '}',
        'exit 0'
    ) -join "`r`n")
}


function New-AppInstallerUninstallScript {
    param([Parameter(Mandatory)][string]$PackageIdentity)

    return (@(
        '$ErrorActionPreference = ''Stop''',
        'try {',
        ('    foreach ($package in @(Get-AppxPackage -AllUsers -Name ''{0}'')) {{' -f $PackageIdentity),
        '        # A profile that is not loaded fails with 0x80073D19; the provisioning is removed anyway.',
        '        try { Remove-AppxPackage -Package $package.PackageFullName -AllUsers -ErrorAction Stop } catch { }',
        '    }',
        ('    foreach ($provisioned in @(Get-AppxProvisionedPackage -Online | Where-Object {{ $_.DisplayName -eq ''{0}'' }})) {{' -f $PackageIdentity),
        '        Remove-AppxProvisionedPackage -Online -PackageName $provisioned.PackageName | Out-Null',
        '    }',
        ('    if (@(Get-AppxProvisionedPackage -Online | Where-Object {{ $_.DisplayName -eq ''{0}'' }}).Count -gt 0) {{ throw ''The provisioned package is still present.'' }}' -f $PackageIdentity),
        '}',
        'catch {',
        '    [Console]::Error.WriteLine($_.Exception.Message)',
        '    exit 1',
        '}',
        'exit 0'
    ) -join "`r`n")
}


function New-AppInstallerDetectionScript {
    <#
    .SYNOPSIS
        Returns the detection script: the provisioned package at the packaged
        version or newer.

    .DESCRIPTION
        Provisioning writes one key per package full name under the all-user
        store, and removing the provisioning removes the key. The script reads
        the 64-bit view, so it needs no elevation and gives the same answer
        from a 32-bit host.
    #>
    param(
        [Parameter(Mandatory)][string]$PackageIdentity,
        [Parameter(Mandatory)][string]$MinimumVersion
    )

    return (@(
        ('$minimum = [version]''{0}''' -f $MinimumVersion),
        '$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)',
        '$store = $base.OpenSubKey(''SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Applications'')',
        'if ($store) {',
        '    foreach ($name in $store.GetSubKeyNames()) {',
        ('        if ($name -match (''^'' + [regex]::Escape(''{0}'') + ''_(?<v>\d+(\.\d+){{1,3}})_[^_]+_[^_]*_8wekyb3d8bbwe$'')) {{' -f $PackageIdentity),
        '            if ([version]$Matches[''v''] -ge $minimum) { Write-Output ''Installed''; break }',
        '        }',
        '    }',
        '}',
        'exit 0'
    ) -join "`r`n")
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageAppInstaller {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "App Installer (WinGet) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $release = Get-WinGetRelease
    $version = $release.Version

    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Download and verify ---
    $localBundle  = Get-VerifiedAsset -Asset $release.Bundle   -Folder $BaseDownloadRoot
    $localDeps    = Get-VerifiedAsset -Asset $release.DepsZip  -Folder $BaseDownloadRoot
    $localJson    = Get-VerifiedAsset -Asset $release.DepsJson -Folder $BaseDownloadRoot
    $localLicense = Get-VerifiedAsset -Asset $release.License  -Folder $BaseDownloadRoot

    $signature = Get-AuthenticodeSignature -LiteralPath $localBundle
    $signerSubject = if ($signature.SignerCertificate) { [string]$signature.SignerCertificate.Subject } else { '' }
    if ($signature.Status -ne 'Valid' -or $signerSubject -notmatch 'CN=Microsoft Corporation') {
        throw "The bundle signature is not a valid Microsoft signature (status: $($signature.Status), signer: $signerSubject)."
    }
    Write-Log "Authenticode signature       : Valid, Microsoft Corporation"

    $identity = Get-BundleIdentity -Path $localBundle
    $packageVersion = [string]$identity.Version
    if ([string]$identity.Name -ne $PackageIdentity) {
        throw "The bundle identity is '$($identity.Name)' but '$PackageIdentity' was expected."
    }
    if ($packageVersion -notmatch '^\d+(\.\d+){3}$') {
        throw "The bundle identity version is not four numbers: '$packageVersion'."
    }
    Write-Log "Bundle identity              : $($identity.Name) $packageVersion"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    Copy-Item -LiteralPath $localBundle  -Destination (Join-Path $localContentPath $release.Bundle.Name)  -Force -ErrorAction Stop
    Copy-Item -LiteralPath $localLicense -Destination (Join-Path $localContentPath $release.License.Name) -Force -ErrorAction Stop
    Write-Log "Copied bundle and license to : $localContentPath"

    $depsFolder = Join-Path $localContentPath "deps"
    Expand-Dependencies -ZipPath $localDeps -JsonPath $localJson -Destination $depsFolder -Architecture $Architectures

    # --- Content wrappers ---
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content   (New-AppInstallerInstallScript -BundleFileName $release.Bundle.Name) `
        -UninstallPs1Content (New-AppInstallerUninstallScript -PackageIdentity $PackageIdentity)

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $AppName
        Publisher       = $Publisher
        SoftwareVersion = $version
        InstallerFile   = $release.Bundle.Name
        InstallerType   = "MSIX"
        InstallArgs     = ""
        UninstallArgs   = ""
        RunningProcess  = @()
        Detection       = @{
            Type           = "Script"
            ScriptLanguage = "PowerShell"
            ScriptText     = (New-AppInstallerDetectionScript -PackageIdentity $PackageIdentity -MinimumVersion $packageVersion)
        }
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Detection                    : provisioned $PackageIdentity version $packageVersion or newer"
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageAppInstaller {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "App Installer (WinGet) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }
    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()

    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection type               : $($manifest.Detection.Type)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    Sync-StagedContentToNetwork -LocalContentPath $localContentPath -NetworkContentPath $networkContentPath -Manifest $manifest

    New-MECMApplicationFromManifest `
        -Manifest $manifest `
        -SiteCode $SiteCode `
        -Comment $Comment `
        -NetworkContentPath $networkContentPath `
        -EstimatedRuntimeMins $EstimatedRuntimeMins `
        -MaximumRuntimeMins $MaximumRuntimeMins
}


# --- Latest-only mode ---
if ($GetLatestVersionOnly) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $release = Get-WinGetRelease -Quiet
        Write-Output $release.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("App Installer GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "App Installer (WinGet) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "GitHubApiUrl                 : $GitHubApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageAppInstaller
    }
    elseif ($PackageOnly) {
        Invoke-PackageAppInstaller
    }
    else {
        Invoke-StageAppInstaller
        Invoke-PackageAppInstaller
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-appinstaller'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
