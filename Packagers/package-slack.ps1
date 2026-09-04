<#
Vendor: Slack Technologies
App: Slack
CMName: Slack
VendorUrl: https://slack.com/
CPE: cpe:2.3:a:slack:slack:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://slack.com/release-notes/windows
DownloadPageUrl: https://slack.com/downloads/windows
IconSource: External
UpdateCadenceDays: 14

.SYNOPSIS
    Packages Slack (x64) MSIX for MECM.

.DESCRIPTION
    Queries the Slack desktop release API for the current Windows x64 MSIX
    version, downloads the signed package, stages content to a versioned
    local folder, and creates an MECM Application that provisions the
    package for every user (Add-AppxProvisionedPackage) with file-existence
    detection on the provisioned package folder. Slack no longer publishes
    the machine-wide MSI; the MSIX is its documented enterprise install.

    Supports two-phase operation:
      -StageOnly    Download, verify identity and signature, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Slack Technologies\Slack\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Slack).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download the MSIX, verify identity and signature,
    generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with file-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Slack version string and exits.

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
    [int]$EstimatedRuntimeMins = 15,
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
# Slack publishes the desktop client for machine-wide deployment as an MSIX;
# the release API still answers the msi variant with a URL the CDN no longer
# serves, so the msix variant is the source of record.
$VersionApiUrl = "https://slack.com/api/desktop.latestRelease?platform=windows&variant=msix&arch=x64"
$DownloadBase  = "https://downloads.slack-edge.com/desktop-releases/windows/x64"
$MsixFileName  = "Slack.msix"
$PackageIdentity = "com.tinyspeck.slackdesktop"

$VendorFolder = "Slack Technologies"
$AppFolder    = "Slack"

$BaseDownloadRoot = Join-Path $DownloadRoot "Slack"

# --- Functions ---


function Assert-MsixPayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the ZIP local-file header.
    .DESCRIPTION
        An MSIX is an OPC (ZIP) container; a CDN error page would otherwise
        stage as a valid-looking package and fail only when the AppX cmdlet
        opens it.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $buffer = New-Object byte[] 4
    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, 4) } finally { $stream.Dispose() }

    if ($read -lt 4 -or $buffer[0] -ne 0x50 -or $buffer[1] -ne 0x4B) {
        throw "Downloaded payload is not an MSIX container (no ZIP header): $Path"
    }
}


function Get-MsixIdentity {
    <#
    .SYNOPSIS
        Returns the Name, Version, Publisher and ProcessorArchitecture from an
        MSIX manifest.
    #>
    param([Parameter(Mandatory)][string]$Path)

    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry('AppxManifest.xml')
        if (-not $entry) { throw "AppxManifest.xml not found in MSIX: $Path" }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        try { [xml]$manifest = $reader.ReadToEnd() } finally { $reader.Dispose() }
    }
    finally { $zip.Dispose() }

    return $manifest.Package.Identity
}


function Get-AppxPublisherId {
    <#
    .SYNOPSIS
        Derives the 13-character publisher id that names a package folder
        under Program Files\WindowsApps.
    .DESCRIPTION
        SHA-256 of the UTF-16LE publisher string, first eight bytes, encoded
        as 13 Douglas Crockford base32 digits (the alphabet omits i, l, o, u),
        the same derivation the AppX runtime uses for PackageFamilyName.
    #>
    param([Parameter(Mandatory)][string]$Publisher)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash([System.Text.Encoding]::Unicode.GetBytes($Publisher)) } finally { $sha.Dispose() }
    $bits = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt 8; $i++) { [void]$bits.Append([Convert]::ToString($hash[$i], 2).PadLeft(8, '0')) }
    [void]$bits.Append('0')
    $alphabet = '0123456789abcdefghjkmnpqrstvwxyz'
    $out = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt 65; $i += 5) {
        [void]$out.Append($alphabet[[Convert]::ToInt32($bits.ToString($i, 5), 2)])
    }
    return $out.ToString()
}


function Get-SlackRelease {
    <#
    .SYNOPSIS
        Returns the current Slack x64 MSIX release as Version and Url.

    .DESCRIPTION
        The API reports the download URL alongside the version; using the
        reported URL keeps the packager working when the vendor changes the
        content path. The documented path template is the fallback when the
        response omits the URL.
    #>
    param([switch]$Quiet)

    Write-Log "Slack version API URL        : $VersionApiUrl" -Quiet:$Quiet

    try {
        $jsonText = (curl.exe -L --fail --silent --show-error $VersionApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Slack release API: $VersionApiUrl" }

        $json = ConvertFrom-Json $jsonText
        $version = [string]$json.version
        if ([string]::IsNullOrWhiteSpace($version)) { throw "version field was empty." }
        if ($version -notmatch '^\d+(\.\d+)+$') { throw "Unexpected version value: $version" }

        $url = [string]$json.download_url
        if ([string]::IsNullOrWhiteSpace($url) -or $url -notmatch '\.msix$') {
            $url = "$DownloadBase/$version/$MsixFileName"
        }

        Write-Log "Latest Slack version         : $version" -Quiet:$Quiet
        return [pscustomobject]@{ Version = $version; Url = $url }
    }
    catch {
        Write-Log "Failed to get Slack release: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageSlack {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Slack (x64 MSIX) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Get version ---
    $release = Get-SlackRelease
    if (-not $release) { throw "Could not resolve Slack version." }

    $version     = $release.Version
    $downloadUrl = $release.Url

    Write-Log "Version                      : $version"
    Write-Log "Package filename             : $MsixFileName"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log ""

    # --- Download ---
    # The CDN publishes every version under the same file name, so the cache
    # is keyed by version folder rather than by file name.
    $versionCache = Join-Path $BaseDownloadRoot ("cache-" + $version)
    Initialize-Folder -Path $versionCache
    $localMsix = Join-Path $versionCache $MsixFileName
    if (-not (Test-Path -LiteralPath $localMsix)) {
        Write-Log "Downloading MSIX..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localMsix
    }
    else {
        Write-Log "Local MSIX exists. Skipping download."
    }

    Assert-MsixPayload -Path $localMsix

    # --- Identity and signature ---
    $identity = Get-MsixIdentity -Path $localMsix
    $identityName    = [string]$identity.Name
    $identityVersion = [string]$identity.Version
    if ($identityName -ne $PackageIdentity) {
        throw "MSIX identity is '$identityName' but '$PackageIdentity' was expected; refusing to stage an unexpected package."
    }
    if ($identity.ProcessorArchitecture -ne 'x64') {
        throw "MSIX architecture is '$($identity.ProcessorArchitecture)' but x64 was expected."
    }
    if (-not $identityVersion.StartsWith($version + '.')) {
        throw "MSIX identity version '$identityVersion' does not match the API version '$version'."
    }

    # The AppX cmdlets refuse an unsigned or untrusted package, so the signing
    # certificate is pinned into the install wrapper at stage time.
    $signature = Get-AuthenticodeSignature -LiteralPath $localMsix
    if ($signature.Status -ne 'Valid') {
        throw "MSIX signature is not valid (status: $($signature.Status))."
    }
    $thumbprint = $signature.SignerCertificate.Thumbprint
    $publisherId = Get-AppxPublisherId -Publisher ([string]$identity.Publisher)

    Write-Log "MSIX identity name           : $identityName"
    Write-Log "MSIX identity version        : $identityVersion"
    Write-Log "MSIX publisher id            : $publisherId"
    Write-Log "MSIX signer thumbprint       : $thumbprint"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedMsix = Join-Path $localContentPath $MsixFileName
    if (-not (Test-Path -LiteralPath $stagedMsix)) {
        Copy-Item -LiteralPath $localMsix -Destination $stagedMsix -Force -ErrorAction Stop
        Write-Log "Copied MSIX to staged folder : $stagedMsix"
    }
    else {
        Write-Log "Staged MSIX exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    # Provisioned for every user (Add-AppxProvisionedPackage -Regions all is
    # the vendor's documented enterprise command); a user who already runs a
    # per-user install receives the provisioned package at next sign-in.
    $installScript = @"
`$msixPath = Join-Path `$PSScriptRoot '$MsixFileName'
if (-not (Test-Path -LiteralPath `$msixPath)) { Write-Error "MSIX not found: `$msixPath"; exit 1 }
`$sig = Get-AuthenticodeSignature -LiteralPath `$msixPath
if (`$sig.Status -ne 'Valid') { Write-Error "MSIX signature not valid: `$(`$sig.Status)"; exit 2 }
if (`$sig.SignerCertificate.Thumbprint -ne '$thumbprint') { Write-Error "MSIX signed by unexpected cert: `$(`$sig.SignerCertificate.Thumbprint)"; exit 3 }
Add-AppxProvisionedPackage -Online -PackagePath `$msixPath -SkipLicense -Regions all | Out-Null
exit 0
"@

    $uninstallScript = @"
`$prov = Get-AppxProvisionedPackage -Online |
    Where-Object { `$_.DisplayName -eq '$identityName' }
foreach (`$p in `$prov) {
    Remove-AppxProvisionedPackage -Online -PackageName `$p.PackageName | Out-Null
}
Get-AppxPackage -AllUsers -Name '$identityName' | ForEach-Object {
    Remove-AppxPackage -Package `$_.PackageFullName -AllUsers
}
exit 0
"@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installScript `
        -UninstallPs1Content $uninstallScript

    # --- Detection ---
    # A provisioned package lands under Program Files\WindowsApps in a folder
    # named from the identity, version, architecture and publisher id; the
    # file clause needs no script signing on the client.
    $packageFolder = "{0}\WindowsApps\{1}_{2}_x64__{3}" -f $env:ProgramFiles, $identityName, $identityVersion, $publisherId

    Write-Log "Detection path               : $packageFolder"
    Write-Log "Detection file               : app\Slack.exe"
    Write-Log ""

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Slack"
        Publisher       = "Slack Technologies"
        SoftwareVersion = $version
        InstallerFile   = $MsixFileName
        InstallerType   = "MSIX"
        InstallArgs     = ""
        UninstallArgs   = ""
        RunningProcess  = @("slack")
        Detection       = @{
            Type         = "File"
            FilePath     = (Join-Path $packageFolder "app")
            FileName     = "Slack.exe"
            PropertyType = "Existence"
            Is64Bit      = $true
        }
    }

    # Save version marker for Package phase
    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageSlack {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Slack (x64 MSIX) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    # --- Resolve version from local staging ---
    Initialize-Folder -Path $BaseDownloadRoot

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }
    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()

    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    # --- Read manifest ---
    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Path               : $($manifest.Detection.FilePath)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
    Write-Log ""

    # --- Network share ---
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    # --- Copy staged content to network ---
    Sync-StagedContentToNetwork -LocalContentPath $localContentPath -NetworkContentPath $networkContentPath -Manifest $manifest

    # --- MECM application ---
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
        $release = Get-SlackRelease -Quiet
        if (-not $release) { exit 1 }
        Write-Output $release.Version
        exit 0
    }
    catch {
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Slack (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "VersionApiUrl                : $VersionApiUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageSlack
    }
    elseif ($PackageOnly) {
        Invoke-PackageSlack
    }
    else {
        Invoke-StageSlack
        Invoke-PackageSlack
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-slack'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
