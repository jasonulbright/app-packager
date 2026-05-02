<#
Vendor: Postman
App: Postman (User)
CMName: Postman
VendorUrl: https://www.postman.com/
CPE: cpe:2.3:a:postman:postman:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.postman.com/release-notes/postman-app/
DownloadPageUrl: https://www.postman.com/downloads/
UpdateCadenceDays: 14

.SYNOPSIS
    Packages the Postman desktop app for MECM user-context deployment.

.DESCRIPTION
    Downloads the public Postman Windows x64 desktop installer, derives the
    version from the installer file metadata, stages content to a versioned
    local folder, and creates a user-context MECM Application with file-version
    detection under LOCALAPPDATA.

    The public Postman Windows installer is a per-user desktop app. The
    authenticated Postman Enterprise MSI is the system-wide and officially
    documented silent deployment path, but that package is available only to
    Postman Enterprise admins inside Postman organization settings.

    Supports two-phase operation:
      -StageOnly    Download EXE, write wrappers and manifest
      -PackageOnly  Read manifest, copy to network, create MECM app

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (Package phase)
    - Write access to FileServerPath (Package phase)
    - Internet access (Stage phase)
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 10,
    [int]$MaximumRuntimeMins = 25,
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

$DownloadUrl = "https://dl.pstmn.io/download/latest/win64"
$InstallerFileName = "Postman-x64-Setup.exe"

$VendorFolder = "Postman"
$AppFolder    = "Postman"
$BaseDownloadRoot = Join-Path $DownloadRoot "Postman"

function Get-PostmanExeVersion {
    param([Parameter(Mandatory)][string]$ExePath)

    $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($ExePath)
    $version = $vi.ProductVersion
    if ([string]::IsNullOrWhiteSpace($version)) { $version = $vi.FileVersion }
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Could not read Postman version from installer metadata: $ExePath"
    }

    $clean = ([string]$version).Trim()
    $m = [regex]::Match($clean, '\d+(?:\.\d+){1,3}')
    if ($m.Success) { return $m.Value }
    return $clean
}

function Get-LatestPostmanRelease {
    param([switch]$Quiet)

    Write-Log "Postman download URL         : $DownloadUrl" -Quiet:$Quiet

    try {
        Initialize-Folder -Path $BaseDownloadRoot

        $localExe = Join-Path $BaseDownloadRoot $InstallerFileName
        Invoke-DownloadWithRetry -Url $DownloadUrl -OutFile $localExe -Quiet:$Quiet

        $version = Get-PostmanExeVersion -ExePath $localExe
        Write-Log "Latest Postman version       : $version" -Quiet:$Quiet

        return [pscustomobject]@{
            Version     = $version
            FileName    = $InstallerFileName
            DownloadUrl = $DownloadUrl
            LocalPath   = $localExe
        }
    }
    catch {
        Write-Log "Failed to get Postman version: $($_.Exception.Message)" -Level ERROR -Quiet:$Quiet
        return $null
    }
}

function New-PostmanInstallWrapper {
    param([Parameter(Mandatory)][string]$InstallerFile)

    # Postman ships a Squirrel-based installer; Postman-x64-Setup.exe -s spawns the
    # actual extract/install in child processes and the parent exits 0 within a few
    # seconds. Without polling for the resulting Postman.exe, the wrapper would
    # report success before the install finishes and CCM would mark the deployment
    # complete against an empty install dir.
    return (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFile),
        'if (-not (Test-Path -LiteralPath $exePath)) { Write-Error "Missing Postman installer"; exit 2 }',
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''-s'') -Wait -PassThru -NoNewWindow',
        '$installedExe = Join-Path $env:LOCALAPPDATA ''Postman\Postman.exe''',
        '$deadline = (Get-Date).AddMinutes(5)',
        'while (-not (Test-Path -LiteralPath $installedExe)) {',
        '    if ((Get-Date) -gt $deadline) {',
        '        Write-Error "Postman install did not produce $installedExe within 5 minutes"',
        '        exit 4',
        '    }',
        '    Start-Sleep -Seconds 5',
        '}',
        'Start-Sleep -Seconds 3',
        'try { Stop-Process -Name Postman -Force -ErrorAction SilentlyContinue } catch { }',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}

function New-PostmanUninstallWrapper {
    return (
        '$updateExe = Join-Path $env:LOCALAPPDATA ''Postman\Update.exe''',
        'if (-not (Test-Path -LiteralPath $updateExe)) { exit 0 }',
        '$proc = Start-Process -FilePath $updateExe -ArgumentList @(''--uninstall'', ''-s'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}

function Invoke-StagePostman {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Postman (User) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $release = Get-LatestPostmanRelease
    if (-not $release) { throw "Could not resolve Postman release info." }

    $version = $release.Version
    $localInstaller = $release.LocalPath

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $DownloadUrl"
    Write-Log "Local installer              : $localInstaller"
    Write-Log "Install context              : User"
    Write-Log ""

    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    Copy-Item -LiteralPath $localInstaller -Destination (Join-Path $localContentPath $InstallerFileName) -Force -ErrorAction Stop
    Write-Log "Copied installer to stage    : $localContentPath"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content (New-PostmanInstallWrapper -InstallerFile $InstallerFileName) `
        -UninstallPs1Content (New-PostmanUninstallWrapper)

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName                  = "Postman (User) - $version"
        Publisher                = "Postman"
        SoftwareVersion          = $version
        DisplayName              = "Postman"
        InstallerFile            = $InstallerFileName
        InstallerType            = "EXE"
        InstallArgs              = "-s"
        UninstallCommand         = "%LOCALAPPDATA%\Postman\Update.exe"
        UninstallArgs            = "--uninstall -s"
        RunningProcess           = @("Postman")
        InstallationBehaviorType = "InstallForUser"
        LogonRequirementType     = "OnlyWhenUserLoggedOn"
        Detection                = @{
            Type          = "File"
            FilePath      = "%LOCALAPPDATA%\Postman"
            FileName      = "Postman.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $false
        }
    }

    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"
    return $localContentPath
}

function Invoke-PackagePostman {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Postman (User) - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    $versionFile = Join-Path $BaseDownloadRoot "staged-version.txt"
    if (-not (Test-Path -LiteralPath $versionFile)) {
        throw "Version marker not found - run Stage phase first: $versionFile"
    }

    $version = (Get-Content -LiteralPath $versionFile -Raw -ErrorAction Stop).Trim()
    $localContentPath = Join-Path $BaseDownloadRoot $version
    $manifest = Read-StageManifest -Path (Join-Path $localContentPath "stage-manifest.json")

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Install behavior             : $($manifest.InstallationBehaviorType)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $manifest.SoftwareVersion
    Initialize-Folder -Path $networkContentPath

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

if ($GetLatestVersionOnly) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $release = Get-LatestPostmanRelease -Quiet
        if (-not $release) { exit 1 }
        Write-Output $release.Version
        exit 0
    }
    catch {
        exit 1
    }
}

try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Postman (User) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StagePostman
    }
    elseif ($PackageOnly) {
        Invoke-PackagePostman
    }
    else {
        Invoke-StagePostman
        Invoke-PackagePostman
    }

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
