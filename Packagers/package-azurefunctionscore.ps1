<#
Vendor: Microsoft
App: Azure Functions Core Tools
CMName: Azure Functions Core Tools
VendorUrl: https://learn.microsoft.com/azure/azure-functions/functions-run-local
CPE: cpe:2.3:a:microsoft:azure_functions_core_tools:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/Azure/azure-functions-core-tools/releases
DownloadPageUrl: https://github.com/Azure/azure-functions-core-tools/releases/latest
UpdateCadenceDays: 30

.SYNOPSIS
    Packages Azure Functions Core Tools (x64) MSI for MECM.

.DESCRIPTION
    Resolves the latest Azure Functions Core Tools release from the GitHub
    releases API, downloads the x64 MSI, stages content to a versioned local
    folder with ARP detection metadata, and creates an MECM Application with
    registry-based detection.

    Supports two-phase operation:
      -StageOnly    Download, derive ARP detection from MSI properties, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\Azure Functions Core Tools\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\Azure Functions Core Tools).
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download installer, derive ARP detection from MSI
    properties, generate content wrappers and stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the latest available Azure Functions Core Tools version string
    and exits.

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
$GitHubApiUrl = "https://api.github.com/repos/Azure/azure-functions-core-tools/releases/latest"

$VendorFolder = "Microsoft"
$AppFolder    = "Azure Functions Core Tools"

$BaseDownloadRoot = Join-Path $DownloadRoot "Azure Functions Core Tools"

# --- Functions ---


function Assert-FuncCliPayloadIsMsi {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the OLE compound-document
        signature that every MSI carries.
    .DESCRIPTION
        A release-asset URL that has been retargeted answers 200 with an HTML
        error body, which would otherwise stage as a valid-looking MSI.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $expected = [byte[]](0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1)
    $buffer = New-Object byte[] $expected.Length

    $stream = [System.IO.File]::OpenRead($Path)
    try { $read = $stream.Read($buffer, 0, $buffer.Length) }
    finally { $stream.Dispose() }

    if ($read -lt $expected.Length) {
        throw "Downloaded payload is too small to be an MSI: $Path"
    }
    for ($i = 0; $i -lt $expected.Length; $i++) {
        if ($buffer[$i] -ne $expected[$i]) {
            throw "Downloaded payload is not an MSI (no OLE compound-document header): $Path"
        }
    }
}


function Get-LatestFuncCliRelease {
    <#
    .SYNOPSIS
        Returns Version, FileName and DownloadUrl for the latest x64 MSI.
    .DESCRIPTION
        The release also carries x86 and arm64 MSIs whose names differ only in
        the trailing architecture token, so the match is anchored on '-x64.msi'.
    #>
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        # The GitHub API rejects requests without a User-Agent header.
        $json = (curl.exe -L --fail --silent --show-error -H "User-Agent: app-packager" $GitHubApiUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query GitHub releases API: $GitHubApiUrl" }

        $release = ConvertFrom-Json $json

        $version = $release.tag_name -replace '^v', ''
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Could not parse version from GitHub release tag."
        }

        $asset = $release.assets |
            Where-Object { $_.name -match '(?i)^func-cli-.*-x64\.msi$' } |
            Select-Object -First 1
        if (-not $asset) { throw "No x64 MSI asset found in release $version." }

        Write-Log "Latest Core Tools version    : $version" -Quiet:$Quiet

        return @{
            Version     = $version
            FileName    = $asset.name
            DownloadUrl = $asset.browser_download_url
        }
    }
    catch {
        Write-Log "Failed to get Azure Functions Core Tools release: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageFuncCli {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Azure Functions Core Tools (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $releaseInfo = Get-LatestFuncCliRelease
    if (-not $releaseInfo) { throw "Could not resolve Azure Functions Core Tools release." }

    $msiFileName = $releaseInfo.FileName
    $downloadUrl = $releaseInfo.DownloadUrl

    Write-Log "Release version              : $($releaseInfo.Version)"
    Write-Log "MSI filename                 : $msiFileName"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log ""

    # --- Download ---
    $localMsi = Join-Path $BaseDownloadRoot $msiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Downloading Core Tools MSI..."
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    Assert-FuncCliPayloadIsMsi -Path $localMsi

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName       = $props["ProductName"]
    $productVersionRaw = $props["ProductVersion"]
    $manufacturer      = $props["Manufacturer"]
    $productCode       = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productVersionRaw)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))       { throw "MSI ProductCode missing." }

    Write-Log "MSI ProductName              : $productName"
    Write-Log "MSI ProductVersion           : $productVersionRaw"
    Write-Log "MSI Manufacturer             : $manufacturer"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log ""

    # --- Versioned local content folder ---
    # The release tag carries the fourth field the MSI ProductVersion drops, so
    # the folder is named for the tag to stay addressable from the API.
    $version = $releaseInfo.Version
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $msiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Derive ARP detection from MSI properties ---
    # For standard MSI installs the ARP uninstall key name is the ProductCode GUID.
    $arpRegistryKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + $productCode

    Write-Log "ARP RegistryKey              : $arpRegistryKey"
    Write-Log "ARP DisplayVersion           : $productVersionRaw"
    Write-Log ""

    # --- Generate content wrappers ---
    $wrapperContent = New-MsiWrapperContent -MsiFileName $msiFileName
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    $publisher = $manufacturer
    if ([string]::IsNullOrWhiteSpace($publisher)) { $publisher = "Microsoft Corporation" }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Azure Functions Core Tools $version"
        Publisher       = $publisher
        SoftwareVersion = $version
        InstallerFile   = $msiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("func")
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $arpRegistryKey
            ValueName           = "DisplayVersion"
            ExpectedValue       = $productVersionRaw
            Is64Bit             = $true
        }
    }

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageFuncCli {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Azure Functions Core Tools (x64) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    # --- Resolve version from local staging ---
    Initialize-Folder -Path $BaseDownloadRoot

    $localMsi = Get-ChildItem -Path $BaseDownloadRoot -Filter "func-cli-*-x64.msi" -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $localMsi) {
        throw "Local MSI not found - run Stage phase first: $BaseDownloadRoot"
    }

    # The staged folder is named for the release tag, which the asset filename
    # carries verbatim.
    $m = [regex]::Match($localMsi.Name, '(?i)^func-cli-(?<ver>.+)-x64\.msi$')
    if (-not $m.Success) { throw "Cannot derive version from cached MSI name: $($localMsi.Name)" }

    $localContentPath = Join-Path $BaseDownloadRoot $m.Groups['ver'].Value
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    # --- Read manifest ---
    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
    Write-Log ""

    # --- Network share ---
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkContentPath = Get-NetworkContentPath -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder -Version $manifest.SoftwareVersion -Layout $ContentLayout

    Write-Log "Network content path         : $networkContentPath"
    Write-Log ""

    # --- Copy staged content to network (recursive: a variant split stages
    # its payload in a subfolder) ---
    $localRoot = (Resolve-Path -LiteralPath $localContentPath).Path
    $localFiles = Get-ChildItem -Path $localContentPath -File -Recurse -ErrorAction Stop
    foreach ($f in $localFiles) {
        if ($f.Name -eq "stage-manifest.json") { continue }
        $relative = $f.FullName.Substring($localRoot.Length).TrimStart('\')
        $dest = Join-Path $networkContentPath $relative
        $destDir = Split-Path -Parent $dest
        if (-not (Test-Path -LiteralPath $destDir)) { New-Item -ItemType Directory -Path $destDir -Force -ErrorAction Stop | Out-Null }
        if (-not (Test-Path -LiteralPath $dest)) {
            Copy-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
            Write-Log "Copied to network            : $relative"
        }
        else {
            Write-Log "Already on network           : $relative"
        }
    }

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
        $releaseInfo = Get-LatestFuncCliRelease -Quiet
        if (-not $releaseInfo) { exit 1 }
        Write-Output $releaseInfo.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("Azure Functions Core Tools GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Azure Functions Core Tools (x64) Auto-Packager starting"
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
        Invoke-StageFuncCli
    }
    elseif ($PackageOnly) {
        Invoke-PackageFuncCli
    }
    else {
        Invoke-StageFuncCli
        Invoke-PackageFuncCli
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-azurefunctionscore'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
