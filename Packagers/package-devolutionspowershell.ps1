<#
Vendor: Devolutions
App: PowerShell 7 (Devolutions)
CMName: PowerShell 7 (Devolutions)
VendorUrl: https://github.com/Devolutions/pwsh-distro
CPE: cpe:2.3:a:microsoft:powershell:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://github.com/Devolutions/pwsh-distro/releases
DownloadPageUrl: https://github.com/Devolutions/pwsh-distro/releases
SecurityAdvisoryUrl: https://github.com/PowerShell/PowerShell/security/advisories
IconSource: Installer
WsusSupport: Yes
ArpName: PowerShell 7-x64 (repackaged by Devolutions)

.SYNOPSIS
    Packages PowerShell 7 (x64) from the Devolutions pwsh-distro releases for ConfigMgr.

.DESCRIPTION
    Downloads the Microsoft-payload MSI of the newest pwsh-distro release that
    publishes one. The MSI holds the unmodified Microsoft binaries in an
    installer that Devolutions signs. It installs to
    C:\Program Files\PowerShell\7 and upgrades an existing Microsoft MSI
    installation of PowerShell 7 in place.

    The asset is chosen from the release asset list, and the download is
    checked against the SHA-256 digest that GitHub publishes for the asset and
    against the Authenticode signature of the MSI.

    Detection reads the file version of pwsh.exe at the install path, greater
    than or equal to the packaged version. A device that a software update
    brought to a newer build, or to the same Microsoft binaries, stays
    compliant, and the MSI product code does not matter.

    Supports two-phase operation:
      -StageOnly    Download, verify, read the MSI, write the manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Devolutions\PowerShell 7 (Devolutions)\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path.
    Default: C:\temp\ap

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: download the MSI, verify it, generate content
    wrappers and the stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with file-version detection.

.PARAMETER GetLatestVersionOnly
    Queries the pwsh-distro releases for the newest release that has the
    Microsoft-payload x64 MSI, outputs its version string, and exits. No
    download or ConfigMgr changes are made.

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
$GitHubApiUrl = "https://api.github.com/repos/Devolutions/pwsh-distro/releases?per_page=30"

$VendorFolder = "Devolutions"
$AppFolder    = "PowerShell 7 (Devolutions)"
$AppName      = "PowerShell 7 (Devolutions)"
$Publisher    = "Devolutions"

$SignerName = "DEVOLUTIONS INC."

$BaseDownloadRoot = Join-Path $DownloadRoot "DevolutionsPowerShell"

# --- Functions ---


function Select-DevolutionsPowerShellAsset {
    <#
    .SYNOPSIS
        Picks the newest release that lists the Microsoft-payload x64 MSI.

    .DESCRIPTION
        Releases that carry only the Devolutions-flavored MSI (a different
        product that installs side by side) never match. A draft, a
        prerelease, and an asset without a SHA-256 digest are skipped, because
        the download cannot be verified without the digest.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Json)

    $releases = @(ConvertFrom-Json $Json | ForEach-Object { $_ })
    foreach ($release in $releases) {
        if ($release.draft -or $release.prerelease) { continue }
        foreach ($asset in @($release.assets)) {
            if ([string]$asset.name -notmatch '^Microsoft\.PowerShell-(?<v>\d+(?:\.\d+){2,3})-win-x64\.msi$') { continue }
            $version = $Matches['v']
            if ([string]$asset.digest -notmatch '^sha256:(?<h>[0-9a-fA-F]{64})$') { continue }
            return [pscustomobject]@{
                Version     = $version
                Tag         = [string]$release.tag_name
                AssetName   = [string]$asset.name
                DownloadUrl = [string]$asset.browser_download_url
                Sha256      = $Matches['h'].ToUpperInvariant()
            }
        }
    }
    throw "No pwsh-distro release lists a Microsoft-payload x64 MSI (Microsoft.PowerShell-<version>-win-x64.msi) with a SHA-256 digest."
}


function Get-DevolutionsPowerShellRelease {
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    $json = (curl.exe -L --fail --silent --show-error @(Get-GitHubApiCurlArgs) $GitHubApiUrl) -join ''
    if ($LASTEXITCODE -ne 0) { throw "Failed to query the pwsh-distro releases API." }

    $release = Select-DevolutionsPowerShellAsset -Json $json
    Write-Log "Latest release with the MSI  : $($release.Tag) ($($release.AssetName))" -Quiet:$Quiet
    return $release
}


function Get-MsiFileVersion {
    <#
    .SYNOPSIS
        Reads the Version column of one file in the MSI File table.
    #>
    param(
        [Parameter(Mandatory)][string]$MsiPath,
        [Parameter(Mandatory)][string]$FileName
    )

    $installer = $null; $db = $null; $view = $null; $record = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $installer, @($MsiPath, 0))
        $sql = "SELECT ``FileName``, ``Version`` FROM ``File``"
        $view = $db.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $db, @($sql))
        $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null
        while ($true) {
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
            if ($null -eq $record) { break }
            $name = [string]$record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
            $longName = ($name -split '\|')[-1]
            if ($longName -ieq $FileName) {
                return [string]$record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 2)
            }
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($record) | Out-Null
            $record = $null
        }
        return $null
    }
    finally {
        foreach ($o in @($record, $view, $db, $installer)) {
            if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) {
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) | Out-Null
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageDevolutionsPowerShell {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PowerShell 7 (Devolutions) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $release = Get-DevolutionsPowerShellRelease
    $MsiFileName = $release.AssetName

    Write-Log "Release version              : $($release.Version)"
    Write-Log "Download URL                 : $($release.DownloadUrl)"
    Write-Log ""

    # --- Download and verify ---
    $localMsi = Join-Path $BaseDownloadRoot $MsiFileName
    Write-Log "Local MSI path               : $localMsi"

    if (Test-Path -LiteralPath $localMsi) {
        $cachedHash = (Get-FileHash -LiteralPath $localMsi -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($cachedHash -ne $release.Sha256) {
            Write-Log "Local MSI does not match the published digest. Downloading again." -Level WARN
            Remove-Item -LiteralPath $localMsi -Force -ErrorAction Stop
        }
    }
    if (-not (Test-Path -LiteralPath $localMsi)) {
        Write-Log "Downloading PowerShell 7 MSI..."
        Invoke-DownloadWithRetry -Url $release.DownloadUrl -OutFile $localMsi
    }
    else {
        Write-Log "Local MSI exists. Skipping download."
    }

    $actualHash = (Get-FileHash -LiteralPath $localMsi -Algorithm SHA256 -ErrorAction Stop).Hash
    if ($actualHash -ne $release.Sha256) {
        Remove-Item -LiteralPath $localMsi -Force -ErrorAction SilentlyContinue
        throw "SHA256 mismatch for $MsiFileName (release digest: $($release.Sha256), file: $actualHash)."
    }
    Write-Log "SHA256 matches the digest    : $actualHash"

    $signature = Get-AuthenticodeSignature -LiteralPath $localMsi
    $signerSubject = if ($signature.SignerCertificate) { [string]$signature.SignerCertificate.Subject } else { '' }
    if ($signature.Status -ne 'Valid' -or $signerSubject -notmatch [regex]::Escape("CN=$SignerName")) {
        throw "The MSI signature is not a valid $SignerName signature (status: $($signature.Status), signer: $signerSubject)."
    }
    Write-Log "Authenticode signature       : Valid, $SignerName"
    Write-Log ""

    # --- Extract MSI properties ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi

    $productName       = $props["ProductName"]
    $productVersionRaw = $props["ProductVersion"]
    $manufacturer      = $props["Manufacturer"]
    $productCode       = $props["ProductCode"]

    if ([string]::IsNullOrWhiteSpace($productVersionRaw)) { throw "MSI ProductVersion missing." }
    if ([string]::IsNullOrWhiteSpace($productCode))       { throw "MSI ProductCode missing." }

    $pwshVersion = Get-MsiFileVersion -MsiPath $localMsi -FileName 'pwsh.exe'
    if ([string]::IsNullOrWhiteSpace($pwshVersion) -or $pwshVersion -notmatch '^\d+(\.\d+){1,3}$') {
        throw "The MSI File table has no usable version for pwsh.exe (read: '$pwshVersion')."
    }

    Write-Log "MSI ProductName              : $productName"
    Write-Log "MSI ProductVersion           : $productVersionRaw"
    Write-Log "MSI Manufacturer             : $manufacturer"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log "pwsh.exe file version        : $pwshVersion"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $productVersionRaw
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $MsiFileName
    if (-not (Test-Path -LiteralPath $stagedMsi)) {
        Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
        Write-Log "Copied MSI to staged folder  : $stagedMsi"
    }
    else {
        Write-Log "Staged MSI exists. Skipping copy."
    }

    # --- Content wrappers ---
    $installPs1 = (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/i'', "`"$msiPath`"", ''/qn'', ''/norestart'', ''ADD_PATH=1'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    $uninstallPs1 = (
        ('$msiPath = Join-Path $PSScriptRoot ''{0}''' -f $MsiFileName),
        '$proc = Start-Process msiexec.exe -ArgumentList @(''/x'', "`"$msiPath`"", ''/qn'', ''/norestart'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # --- Write stage manifest ---
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    $manifestData = @{
        AppName         = $AppName
        Publisher       = $Publisher
        SoftwareVersion = $productVersionRaw
        InstallerFile   = $MsiFileName
        InstallerType   = "MSI"
        InstallArgs     = "/qn /norestart ADD_PATH=1"
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("pwsh")
        Detection       = @{
            Type          = "File"
            FilePath      = ("{0}\PowerShell\7" -f $env:ProgramFiles)
            FileName      = "pwsh.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $pwshVersion
            Is64Bit       = $true
        }
    }
    Write-StageManifest -Path $manifestPath -ManifestData $manifestData

    Write-Log ""
    Write-Log "Detection                    : pwsh.exe file version >= $pwshVersion"
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageDevolutionsPowerShell {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "PowerShell 7 (Devolutions) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $msiFiles = Get-ChildItem -Path $BaseDownloadRoot -Filter "Microsoft.PowerShell-*-win-x64.msi" -File
    if (-not $msiFiles -or $msiFiles.Count -eq 0) {
        throw "No staged PowerShell 7 MSI found - run Stage phase first."
    }
    $localMsi = $msiFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    $props = Get-MsiPropertyMap -MsiPath $localMsi.FullName
    if (-not $props -or [string]::IsNullOrWhiteSpace($props["ProductVersion"])) {
        throw "Cannot read ProductVersion from cached MSI."
    }

    $productVersion   = $props["ProductVersion"]
    $localContentPath = Join-Path $BaseDownloadRoot $productVersion
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection File               : $($manifest.Detection.FilePath)\$($manifest.Detection.FileName)"
    Write-Log "Detection Value              : $($manifest.Detection.ExpectedValue)"
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
        $release = Get-DevolutionsPowerShellRelease -Quiet
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
    Write-Log "PowerShell 7 (Devolutions) Auto-Packager starting"
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
        Invoke-StageDevolutionsPowerShell
    }
    elseif ($PackageOnly) {
        Invoke-PackageDevolutionsPowerShell
    }
    else {
        Invoke-StageDevolutionsPowerShell
        Invoke-PackageDevolutionsPowerShell
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-devolutionspowershell'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
