<#
Vendor: Microsoft
App: PowerShell 7 (MSIX)
CMName: PowerShell 7 (MSIX)
VendorUrl: https://github.com/PowerShell/PowerShell
CPE: cpe:2.3:a:microsoft:powershell:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/PowerShell/PowerShell/releases
DownloadPageUrl: https://github.com/PowerShell/PowerShell/releases
SecurityAdvisoryUrl: https://github.com/PowerShell/PowerShell/security/advisories
IconSource: External
WsusSupport: No (InstallerTypeUnsupported|CustomInstall)
ArpName: none

.SYNOPSIS
    Packages PowerShell 7 as the MSIX package for ConfigMgr.

.DESCRIPTION
    Downloads the newest stable PowerShell release that carries the msixbundle
    from the PowerShell GitHub releases, checks the file against the SHA-256
    digest that GitHub publishes for the asset, and requires a valid Microsoft
    signature and the identity Microsoft.PowerShell. From PowerShell 7.7.0 the
    project publishes no MSI, so the MSIX package is the Microsoft package for
    a computer that needs the current release.

    The install script provisions the bundle for all users with
    Add-AppxProvisionedPackage. The package has no dependency packages. It
    needs no Microsoft Store and no network access. Provisioning reaches an
    existing user profile at the next sign-in; the script does not register the
    package for a user who is signed in.

    Limits of the MSIX package (Microsoft Learn, "Install PowerShell 7 on
    Windows", Limitations of a MSIX-based installation):
      - It is a single-user installation. There is no option to install it
        for all users; provisioning gives every profile its own copy.
      - It does not support PowerShell remoting. The application sandbox blocks
        changes to the application root folder, so the WSMan configuration
        cannot be changed and remote sessions cannot connect to it.
      - It does not support Register-PSSessionConfiguration,
        Update-Help -Scope AllUsers, Enable-ExperimentalFeature -Scope AllUsers
        or Set-ExecutionPolicy -Scope LocalMachine, and it cannot create the
        all-users profile scripts. Current-user profiles work.
      - It installs below Program Files\WindowsApps, not to
        Program Files\PowerShell\7, and it does not update through Microsoft
        Update. The MSI-based PowerShell 7 packager is the choice for a server
        or for a computer that needs remoting or the all-users scope.

    Detection needs the provisioning in the all-user store and accepts the
    packaged version of the x64 package folder below Program Files\WindowsApps
    or a newer one. The application version is the release version of the
    package (for example 7.7.0). The provisioning key carries the date-based
    bundle identity version (2026.1110.418.0), which cannot be compared with
    the release version, so the folder version decides.

    The uninstall script removes the per-user copies of signed-in users and
    the provisioning. It does not remove the copy of a logged-off profile,
    because that profile is not loaded.

    Supports two-phase operation:
      -StageOnly    Download, verify, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (for example, "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the ConfigMgr application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\PowerShell 7 (MSIX)\<Version>

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
    Queries the PowerShell GitHub releases, outputs the version of the newest
    stable release that carries the msixbundle, and exits. No download or
    ConfigMgr changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - Configuration Manager console installed (ConfigurationManager PowerShell module available)
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
$GitHubApiUrl = "https://api.github.com/repos/PowerShell/PowerShell/releases?per_page=30"

$VendorFolder = "Microsoft"
$AppFolder    = "PowerShell 7 (MSIX)"
$AppName      = "PowerShell 7 (MSIX)"
$Publisher    = "Microsoft Corporation"

$PackageIdentity = "Microsoft.PowerShell"

$BaseDownloadRoot = Join-Path $DownloadRoot "PowerShell7Msix"

# --- Functions ---


function Select-PowerShellMsixRelease {
    <#
    .SYNOPSIS
        Picks the newest stable release that lists the msixbundle with a digest.

    .DESCRIPTION
        A draft, a prerelease and a tag with a suffix (v7.7.0-rc.1) never
        match. The LTS bundle (PowerShell-LTS-x.y.z.msixbundle) and the preview
        bundle (PowerShellPreview-...) are other packages and never match. An
        asset without a SHA-256 digest is not used, because its download cannot
        be verified.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Json)

    $releases = @(ConvertFrom-Json $Json | ForEach-Object { $_ })
    foreach ($release in $releases) {
        if ($release.draft -or $release.prerelease) { continue }
        $tag = [string]$release.tag_name
        if ($tag -notmatch '^v\d+(?:\.\d+){2}$') { continue }

        foreach ($asset in @($release.assets)) {
            if ([string]$asset.name -notmatch '^PowerShell-(?<v>\d+(?:\.\d+){2})\.msixbundle$') { continue }
            $version = $Matches['v']
            if ([string]$asset.digest -notmatch '^sha256:(?<h>[0-9a-fA-F]{64})$') { continue }
            return [pscustomobject]@{
                Version     = $version
                Tag         = $tag
                Name        = [string]$asset.name
                DownloadUrl = [string]$asset.browser_download_url
                Sha256      = $Matches['h'].ToUpperInvariant()
            }
        }
    }
    throw "No stable PowerShell release lists the msixbundle with a SHA-256 digest."
}


function Get-PowerShellMsixRelease {
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    $json = (curl.exe -L --fail --silent --show-error @(Get-GitHubApiCurlArgs) $GitHubApiUrl) -join ''
    if ($LASTEXITCODE -ne 0) { throw "Failed to query the PowerShell releases API." }

    $release = Select-PowerShellMsixRelease -Json $json
    Write-Log "Latest stable release        : $($release.Tag)" -Quiet:$Quiet
    return $release
}


function Get-BundleInfo {
    <#
    .SYNOPSIS
        Returns the identity name and the x64 package version of a bundle.

    .DESCRIPTION
        AppxBundleManifest.xml names the bundle identity (a date-based
        version) and one Package element for each architecture; the Package
        version is the application version that the provisioned package
        carries.
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

    $x64 = @($manifest.Bundle.Packages.Package | Where-Object { $_.Type -eq 'application' -and $_.Architecture -eq 'x64' }) | Select-Object -First 1
    if (-not $x64) { throw "The bundle has no x64 application package." }

    return [pscustomobject]@{
        Name           = [string]$manifest.Bundle.Identity.Name
        BundleVersion  = [string]$manifest.Bundle.Identity.Version
        PackageVersion = [string]$x64.Version
    }
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


function New-PowerShellMsixInstallScript {
    param([Parameter(Mandatory)][string]$BundleFileName)

    return (@(
        '$ErrorActionPreference = ''Stop''',
        'try {',
        '    if ([Environment]::OSVersion.Version -lt [version]''10.0.17763'') { throw ''PowerShell 7 (MSIX) needs Windows 10 version 1809 (build 17763) or later.'' }',
        ('    $bundle = Join-Path $PSScriptRoot ''{0}''' -f $BundleFileName),
        '    if (-not (Test-Path -LiteralPath $bundle)) { throw "The bundle is missing from the content folder: $bundle" }',
        '    Add-AppxProvisionedPackage -Online -PackagePath $bundle -SkipLicense | Out-Null',
        '}',
        'catch {',
        '    [Console]::Error.WriteLine($_.Exception.Message)',
        '    exit 1',
        '}',
        'exit 0'
    ) -join "`r`n")
}


function New-PowerShellMsixUninstallScript {
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


function New-PowerShellMsixDetectionScript {
    <#
    .SYNOPSIS
        Returns the detection script: the package is provisioned and its x64
        application package is at the packaged version or newer.

    .DESCRIPTION
        Provisioning writes one key per bundle under the all-user store, named
        with the date-based bundle identity version (for example
        Microsoft.PowerShell_2026.1110.418.0_neutral_~_8wekyb3d8bbwe), and
        removing the provisioning removes the key. That version cannot be
        compared with the release, so the script also reads the version of
        the x64 package folder below Program Files\WindowsApps
        (Microsoft.PowerShell_7.7.0.0_x64__8wekyb3d8bbwe) and compares that
        one. The script reads the 64-bit registry view and the 64-bit
        Program Files path, so it gives the same answer from a 32-bit host.
    #>
    param(
        [Parameter(Mandatory)][string]$PackageIdentity,
        [Parameter(Mandatory)][string]$MinimumVersion
    )

    return (@(
        ('$minimum = [version]''{0}''' -f $MinimumVersion),
        '$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)',
        '$store = $base.OpenSubKey(''SOFTWARE\Microsoft\Windows\CurrentVersion\Appx\AppxAllUserStore\Applications'')',
        '$provisioned = $false',
        'if ($store) {',
        '    foreach ($name in $store.GetSubKeyNames()) {',
        ('        if ($name -match (''^'' + [regex]::Escape(''{0}'') + ''_\d+(\.\d+){{1,3}}_neutral_~_8wekyb3d8bbwe$'')) {{ $provisioned = $true; break }}' -f $PackageIdentity),
        '    }',
        '}',
        'if ($provisioned) {',
        '    $apps = Join-Path $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }) ''WindowsApps''',
        '    foreach ($dir in @(Get-ChildItem -LiteralPath $apps -Directory -ErrorAction SilentlyContinue)) {',
        ('        if ($dir.Name -match (''^'' + [regex]::Escape(''{0}'') + ''_(?<v>\d+(\.\d+){{1,3}})_x64__8wekyb3d8bbwe$'')) {{' -f $PackageIdentity),
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

function Invoke-StagePowerShellMsix {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PowerShell 7 (MSIX) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $release = Get-PowerShellMsixRelease
    $version = $release.Version

    Write-Log "Version                      : $version"
    Write-Log ""

    # --- Download and verify ---
    $localBundle = Get-VerifiedAsset -Asset $release -Folder $BaseDownloadRoot

    $signature = Get-AuthenticodeSignature -LiteralPath $localBundle
    $signerSubject = if ($signature.SignerCertificate) { [string]$signature.SignerCertificate.Subject } else { '' }
    if ($signature.Status -ne 'Valid' -or $signerSubject -notmatch 'CN=Microsoft Corporation') {
        throw "The bundle signature is not a valid Microsoft signature (status: $($signature.Status), signer: $signerSubject)."
    }
    Write-Log "Authenticode signature       : Valid, Microsoft Corporation"

    $info = Get-BundleInfo -Path $localBundle
    if ($info.Name -ne $PackageIdentity) {
        throw "The bundle identity is '$($info.Name)' but '$PackageIdentity' was expected."
    }
    if ($info.PackageVersion -notmatch '^\d+(\.\d+){3}$') {
        throw "The package version is not four numbers: '$($info.PackageVersion)'."
    }
    if (-not $info.PackageVersion.StartsWith($version + '.')) {
        throw "The package version '$($info.PackageVersion)' does not match the release version '$version'."
    }
    Write-Log "Bundle identity              : $($info.Name) $($info.BundleVersion)"
    Write-Log "x64 package version          : $($info.PackageVersion)"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    Copy-Item -LiteralPath $localBundle -Destination (Join-Path $localContentPath $release.Name) -Force -ErrorAction Stop
    Write-Log "Copied bundle to             : $localContentPath"

    # --- Content wrappers ---
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content   (New-PowerShellMsixInstallScript -BundleFileName $release.Name) `
        -UninstallPs1Content (New-PowerShellMsixUninstallScript -PackageIdentity $PackageIdentity)

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = $AppName
        Publisher       = $Publisher
        SoftwareVersion = $version
        InstallerFile   = $release.Name
        InstallerType   = "MSIX"
        InstallArgs     = ""
        UninstallArgs   = ""
        RunningProcess  = @("pwsh")
        Detection       = @{
            Type           = "Script"
            ScriptLanguage = "PowerShell"
            ScriptText     = (New-PowerShellMsixDetectionScript -PackageIdentity $PackageIdentity -MinimumVersion $info.PackageVersion)
        }
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Detection                    : provisioned $PackageIdentity x64 package version $($info.PackageVersion) or newer"
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackagePowerShellMsix {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PowerShell 7 (MSIX) - PACKAGE phase"
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
        $release = Get-PowerShellMsixRelease -Quiet
        Write-Output $release.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("PowerShell 7 (MSIX) GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PowerShell 7 (MSIX) Auto-Packager starting"
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
        Invoke-StagePowerShellMsix
    }
    elseif ($PackageOnly) {
        Invoke-PackagePowerShellMsix
    }
    else {
        Invoke-StagePowerShellMsix
        Invoke-PackagePowerShellMsix
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-powershell7msix'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
