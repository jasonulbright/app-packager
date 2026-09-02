<#
Vendor: Jan De Dobbeleer
App: Oh My Posh
CMName: Oh My Posh
VendorUrl: https://ohmyposh.dev/
CPE: cpe:2.3:a:ohmyposh:oh-my-posh:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/JanDeDobbeleer/oh-my-posh/releases
DownloadPageUrl: https://github.com/JanDeDobbeleer/oh-my-posh/releases
UpdateCadenceDays: 30

.SYNOPSIS
    Packages Oh My Posh (x64 MSIX, device-wide) for MECM.

.DESCRIPTION
    Downloads the signed x64 MSIX from the oh-my-posh GitHub releases, stages
    content to a versioned local folder, and creates an MECM Application that
    provisions the package for every user on the device.

    The vendor retired the Inno Setup installer; install-x64.msix is the only
    Windows installer asset the project publishes, so this packager deploys
    through Add-AppxProvisionedPackage rather than a silent EXE.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers. Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes. Default: 10

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes. Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase.

.PARAMETER PackageOnly
    Runs only the Package phase.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Oh My Posh version string and exits.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed
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
$GitHubApiUrl = "https://api.github.com/repos/JanDeDobbeleer/oh-my-posh/releases/latest"

$VendorFolder = "Jan De Dobbeleer"
$AppFolder    = "Oh My Posh"

$BaseDownloadRoot = Join-Path $DownloadRoot "OhMyPosh"
$MsixFileName     = "install-x64.msix"
$PackageIdentity  = "ohmyposh.cli"

# --- Functions ---


function Assert-MsixPayload {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the ZIP local-file header.
    .DESCRIPTION
        An MSIX is an OPC (ZIP) container; a release-asset redirect that lands
        on an HTML error page would otherwise stage as a valid-looking package
        and fail only when the AppX cmdlet opens it.
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
        Returns the Name, Version and ProcessorArchitecture from an MSIX manifest.
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


function Get-LatestOhMyPoshRelease {
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        $json = (curl.exe -L --fail --silent --show-error $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query GitHub releases API." }

        $release = ConvertFrom-Json $json
        $version = [string]$release.tag_name -replace '^v', ''
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Could not parse version from GitHub release tag."
        }

        $asset = $release.assets | Where-Object { $_.name -eq $MsixFileName } | Select-Object -First 1
        if (-not $asset) { throw "No $MsixFileName asset found in release." }

        Write-Log "Latest Oh My Posh version    : $version" -Quiet:$Quiet
        return @{
            Version     = $version
            FileName    = $asset.name
            DownloadUrl = $asset.browser_download_url
        }
    }
    catch {
        Write-Log "Failed to get Oh My Posh version: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageOhMyPosh {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Oh My Posh (x64 MSIX) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestOhMyPoshRelease
    if (-not $releaseInfo) { throw "Could not resolve Oh My Posh version." }

    $version = $releaseInfo.Version

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $($releaseInfo.DownloadUrl)"
    Write-Log "Package filename             : $MsixFileName"
    Write-Log ""

    # --- Download ---
    # The asset name is the same in every release, so the cached copy is
    # replaced rather than reused.
    $localMsix = Join-Path $BaseDownloadRoot $MsixFileName
    Write-Log "Local package path           : $localMsix"
    Write-Log "Downloading Oh My Posh..."
    Invoke-DownloadWithRetry -Url $releaseInfo.DownloadUrl -OutFile $localMsix

    Assert-MsixPayload -Path $localMsix

    # --- Read package identity ---
    $identity = Get-MsixIdentity -Path $localMsix
    $identityName    = [string]$identity.Name
    $identityVersion = [string]$identity.Version

    if ($identityName -ne $PackageIdentity) {
        throw "MSIX identity is '$identityName' but '$PackageIdentity' was expected; refusing to stage an unexpected package."
    }
    if ([string]$identity.ProcessorArchitecture -ne 'x64') {
        throw "MSIX architecture is '$($identity.ProcessorArchitecture)' but x64 was expected."
    }
    if ($identityVersion -notlike "$version.*" -and $identityVersion -ne $version) {
        throw "MSIX identity version '$identityVersion' does not match the release tag '$version'."
    }

    # The AppX cmdlets refuse an unsigned or untrusted package, so the signing
    # certificate is pinned into the install wrapper at stage time.
    $signature = Get-AuthenticodeSignature -LiteralPath $localMsix
    if ($signature.Status -ne 'Valid') {
        throw "MSIX signature is not valid (status: $($signature.Status))."
    }
    $thumbprint = $signature.SignerCertificate.Thumbprint

    Write-Log "MSIX identity name           : $identityName"
    Write-Log "MSIX identity version        : $identityVersion"
    Write-Log "MSIX signer thumbprint       : $thumbprint"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedMsix = Join-Path $localContentPath $MsixFileName
    Copy-Item -LiteralPath $localMsix -Destination $stagedMsix -Force -ErrorAction Stop
    Write-Log "Copied MSIX to staged folder : $stagedMsix"

    # --- Generate content wrappers ---
    $wrapperContent = New-MsixWrapperContent -MsixFileName $MsixFileName -Provisioned $true -SignatureSha1 $thumbprint

    # Removal targets the identity recorded at stage time, so the uninstaller
    # does not depend on reading the manifest back out of the content folder.
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
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $uninstallScript

    # --- Detection ---
    # A provisioned MSIX lands under Program Files\WindowsApps in a directory
    # named with the publisher hash, so detection queries the provisioning
    # store instead of a path.
    $detectionScript = @"
`$wanted = [version]'$identityVersion'
`$prov = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
    Where-Object { `$_.DisplayName -eq '$identityName' }
foreach (`$p in `$prov) {
    `$found = `$null
    if (-not [version]::TryParse([string]`$p.Version, [ref]`$found)) { continue }
    if (`$found -ge `$wanted) { Write-Output 'Installed'; exit 0 }
}
exit 0
"@

    Write-Log "Detection                    : provisioned package '$identityName' with Version >= $identityVersion"
    Write-Log ""

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Oh My Posh $version"
        Publisher       = "Jan De Dobbeleer"
        SoftwareVersion = $version
        InstallerFile   = $MsixFileName
        InstallerType   = "MSIX"
        InstallArgs     = ""
        UninstallArgs   = ""
        RunningProcess  = @("oh-my-posh")
        Detection       = @{
            Type           = "Script"
            ScriptLanguage = "PowerShell"
            ScriptText     = $detectionScript
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

function Invoke-PackageOhMyPosh {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Oh My Posh (x64 MSIX) - PACKAGE phase"
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

    $localFiles = Get-ChildItem -Path $localContentPath -File -ErrorAction Stop
    foreach ($f in $localFiles) {
        if ($f.Name -eq "stage-manifest.json") { continue }
        $dest = Join-Path $networkContentPath $f.Name
        if (-not (Test-Path -LiteralPath $dest)) {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
            Write-Log "Copied to network            : $($f.Name)"
        }
        else {
            Write-Log "Already on network           : $($f.Name)"
        }
    }

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
        $info = Get-LatestOhMyPoshRelease -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Oh My Posh GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Oh My Posh (x64 MSIX) Auto-Packager starting"
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
        Invoke-StageOhMyPosh
    }
    elseif ($PackageOnly) {
        Invoke-PackageOhMyPosh
    }
    else {
        Invoke-StageOhMyPosh
        Invoke-PackageOhMyPosh
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-ohmyposh'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
