<#
Vendor: TODO
App: TODO
CMName: TODO
VendorUrl: https://TODO
CPE: cpe:2.3:a:TODO:TODO:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://TODO
DownloadPageUrl: https://TODO
UpdateCadenceDays: 30

.SYNOPSIS
    Packages TODO (x64) EXE for MECM.

.DESCRIPTION
    Starter template for an EXE-based packager. EXE installers do not
    expose MSI metadata directly, so detection usually falls back to:
      - a file-version check on a well-known binary path, or
      - a registry key written by the installer at install time.

    Read Samples/AUTHORING.md before editing this template.

    Before writing detection code: install the EXE once into a
    disposable VM and confirm what it actually writes. EXE behavior is
    vendor-specific and changes between versions more often than MSI.
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 15,
    [int]$MaximumRuntimeMins = 30,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# ---------------------------------------------------------------------------
# Configuration - edit these per app
# ---------------------------------------------------------------------------

# Latest-version source. This template shows the GitHub Releases API pattern.
# If the vendor uses a normal download page, swap Resolve-LatestRelease for a
# curl + regex scrape (see Samples/package-template-msi.ps1 Resolve-MsiUrl).
$GitHubApiUrl = "https://api.github.com/repos/TODO-owner/TODO-repo/releases/latest"

$VendorFolder = "TODO"
$AppFolder    = "TODO"

$BaseDownloadRoot = Join-Path $DownloadRoot $AppFolder


# ---------------------------------------------------------------------------
# Resolve latest release metadata
# ---------------------------------------------------------------------------
# Returns a hashtable with Version, DownloadUrl, InstallerFileName.
# -Quiet suppresses logging (used by -GetLatestVersionOnly).

function Resolve-LatestRelease {
    param([switch]$Quiet)

    Write-Log "GitHub releases API          : $GitHubApiUrl" -Quiet:$Quiet

    try {
        $response = Invoke-RestMethod -Uri $GitHubApiUrl -UserAgent "AppPackager" -ErrorAction Stop

        # Strip common version prefixes (v1.2.3, desktop-v2026.3.1, etc.)
        $tag = [string]$response.tag_name
        $ver = ($tag -replace '^(v|desktop-v|release-)', '')

        # Find the x64 EXE asset. Adjust the pattern to match the vendor's naming.
        $asset = $response.assets |
            Where-Object { $_.name -match 'TODO.*x64.*\.exe$' } |
            Select-Object -First 1

        if (-not $asset) { throw "No matching x64 .exe asset found in release $tag" }

        Write-Log "Latest version               : $ver" -Quiet:$Quiet
        Write-Log "Download URL                 : $($asset.browser_download_url)" -Quiet:$Quiet

        return @{
            Version           = $ver
            DownloadUrl       = [string]$asset.browser_download_url
            InstallerFileName = [string]$asset.name
        }
    }
    catch {
        Write-Log ("Failed to resolve latest release: " + $_.Exception.Message) -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageApp {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TODO - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $release = Resolve-LatestRelease
    if (-not $release) { throw "Could not resolve latest release." }

    $version       = $release.Version
    $installerName = $release.InstallerFileName
    $downloadUrl   = $release.DownloadUrl

    $localInstaller = Join-Path $BaseDownloadRoot $installerName
    Write-Log "Local installer path         : $localInstaller"
    Write-Log ""
    Write-Log "Downloading installer..."
    Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localInstaller

    # Versioned content folder
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedInstaller = Join-Path $localContentPath $installerName
    if (-not (Test-Path -LiteralPath $stagedInstaller)) {
        Copy-Item -LiteralPath $localInstaller -Destination $stagedInstaller -Force -ErrorAction Stop
        Write-Log "Copied EXE to staged folder  : $stagedInstaller"
    }
    else {
        Write-Log "Staged EXE exists. Skipping copy."
    }

    # -----------------------------------------------------------------
    # Detection: File-version on a known binary path
    # -----------------------------------------------------------------
    # Most EXE installers drop a main binary in C:\Program Files\<App>\.
    # If the binary has a version resource (most do) we can detect on
    # "file at this path with FileVersion matching expected".
    #
    # Install the EXE once on a test VM to confirm:
    #   - the exact install path
    #   - the FileVersion exposed (often differs from marketing version)
    #
    # For apps that only write a registry key, see the Registry block below.

    $detectionPath = "C:\Program Files\TODO"
    $detectionFile = "TODO.exe"
    $detectionType = "File"

    Write-Log "Detection path               : $detectionPath"
    Write-Log "Detection file               : $detectionFile"

    # -----------------------------------------------------------------
    # Install command
    # -----------------------------------------------------------------
    # Most EXE installers support /S or /silent or /quiet or /VERYSILENT.
    # Test on a VM; note the exact switch the vendor documents.
    # For uninstall, you typically read ARP's UninstallString at runtime
    # OR hardcode if the installer is consistent.

    $installArgs   = "/S"                                          # TODO: confirm silent switch
    $uninstallCmd  = "C:\Program Files\TODO\uninstall.exe /S"      # TODO: confirm uninstall path
    $uninstallArgs = ""

    $installPs1 = @"
`$proc = Start-Process -FilePath "`$PSScriptRoot\$installerName" -ArgumentList "$installArgs" -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    $uninstallPs1 = @"
`$proc = Start-Process -FilePath "$uninstallCmd" -ArgumentList "$uninstallArgs" -Wait -PassThru -NoNewWindow
exit `$proc.ExitCode
"@

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    # Stage manifest
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName          = "TODO"                     # display name used for MECM app
        Publisher        = "TODO"
        SoftwareVersion  = $version
        InstallerFile    = $installerName
        InstallerType    = "EXE"
        InstallArgs      = $installArgs
        UninstallCommand = $uninstallCmd
        UninstallArgs    = $uninstallArgs
        RunningProcess   = @()                        # TODO: exe names so MECM can close them before upgrade
        Detection        = @{
            Type     = $detectionType
            Path     = $detectionPath
            FileName = $detectionFile
            Version  = $version                       # or ConvertTo-FileVersion $version if it differs
            Operator = "GreaterEquals"
        }
    }

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"
    return $localContentPath
}


# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------
# Identical flow to the MSI template; Package only cares about the stage
# manifest and the network share.

function Invoke-PackageApp {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TODO - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    # Read the staged manifest. Without running Stage first this will fail;
    # in that case tell the user to Stage before Package.
    $release = Resolve-LatestRelease -Quiet
    if (-not $release) { throw "Could not resolve latest release for manifest lookup." }

    $localContentPath = Join-Path $BaseDownloadRoot $release.Version
    $manifestPath     = Join-Path $localContentPath "stage-manifest.json"

    if (-not (Test-Path -LiteralPath $manifestPath)) {
        throw "Stage manifest not found - run Stage phase first: $manifestPath"
    }

    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Path               : $($manifest.Detection.Path)"
    Write-Log "Detection File               : $($manifest.Detection.FileName)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot     = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

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


# ---------------------------------------------------------------------------
# -GetLatestVersionOnly mode
# ---------------------------------------------------------------------------

if ($GetLatestVersionOnly) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $release = Resolve-LatestRelease -Quiet
        if (-not $release) { exit 1 }
        Write-Output $release.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TODO Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "GitHubApiUrl                 : $GitHubApiUrl"
    Write-Log ""

    if ($StageOnly)       { Invoke-StageApp }
    elseif ($PackageOnly) { Invoke-PackageApp }
    else                  { Invoke-StageApp; Invoke-PackageApp }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
