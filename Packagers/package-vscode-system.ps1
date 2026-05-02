<#
Vendor: Microsoft
App: Visual Studio Code (System)
CMName: Visual Studio Code (System)
VendorUrl: https://code.visualstudio.com/docs/setup/windows
CPE: cpe:2.3:a:microsoft:visual_studio_code:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://code.visualstudio.com/updates
DownloadPageUrl: https://code.visualstudio.com/Download
UpdateCadenceDays: 14

.SYNOPSIS
    Packages Visual Studio Code x64 system installer for MECM.

.DESCRIPTION
    Downloads the latest stable VS Code x64 system installer from the official
    update.code.visualstudio.com endpoint, stages content to a versioned local
    folder, and creates a machine-context MECM Application with file-version
    detection under Program Files.

    This packager is intentionally separate from the user installer package.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager module available)
    - RBAC permissions to create Applications and Deployment Types
    - Local administrator
    - Write access to FileServerPath
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 10,
    [int]$MaximumRuntimeMins = 20,
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

$ReleasesApiUrl = "https://update.code.visualstudio.com/api/releases/stable"
$Quality        = "stable"
$Platform       = "win32-x64"

$VendorFolder = "Microsoft"
$AppFolder    = "Visual Studio Code (System)"
$BaseDownloadRoot = Join-Path $DownloadRoot "VSCodeSystem"

function Get-LatestVSCodeVersion {
    param([switch]$Quiet)

    Write-Log "VS Code releases API         : $ReleasesApiUrl" -Quiet:$Quiet

    try {
        $json = $null
        $lastErr = ''
        foreach ($delay in 0, 1, 2, 4) {
            if ($delay -gt 0) { Start-Sleep -Seconds $delay }
            $out = curl.exe -L --fail --silent --show-error --retry 0 $ReleasesApiUrl 2>&1
            if ($LASTEXITCODE -eq 0) { $json = ($out -join ''); break }
            $lastErr = ($out -join ' ').Trim()
        }
        if (-not $json) { throw "Failed to query VS Code releases API: $lastErr" }

        $versions = ConvertFrom-Json $json
        $version = [string]($versions | Select-Object -First 1)
        if ([string]::IsNullOrWhiteSpace($version)) {
            throw "Could not parse latest VS Code version."
        }

        Write-Log "Latest VS Code version       : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        Write-Log "Failed to get VS Code version: $($_.Exception.Message)" -Level ERROR -Quiet:$Quiet
        return $null
    }
}

function New-VSCodeSystemInstallWrapper {
    param([Parameter(Mandatory)][string]$InstallerFile)

    return (
        ('$installer = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFile),
        'if (-not (Test-Path -LiteralPath $installer)) { Write-Error "Missing VS Code installer"; exit 2 }',
        '$args = @(''/VERYSILENT'', ''/NORESTART'', ''/MERGETASKS=!runcode'')',
        '$proc = Start-Process -FilePath $installer -ArgumentList $args -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}

function New-VSCodeSystemUninstallWrapper {
    return (
        '$uninstall = Join-Path $env:ProgramFiles ''Microsoft VS Code\unins000.exe''',
        'if (-not (Test-Path -LiteralPath $uninstall)) { exit 0 }',
        '$proc = Start-Process -FilePath $uninstall -ArgumentList @(''/VERYSILENT'', ''/NORESTART'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}

function Invoke-StageVSCodeSystem {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Visual Studio Code (System) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $version = Get-LatestVSCodeVersion
    if (-not $version) { throw "Could not resolve VS Code version." }

    $downloadUrl = "https://update.code.visualstudio.com/$version/$Platform/$Quality"
    $installerFileName = "VSCodeSetup-x64-$version.exe"
    $localInstaller = Join-Path $BaseDownloadRoot $installerFileName

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Local installer              : $localInstaller"
    Write-Log ""

    Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localInstaller

    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    Copy-Item -LiteralPath $localInstaller -Destination (Join-Path $localContentPath $installerFileName) -Force -ErrorAction Stop
    Write-Log "Copied installer to stage    : $localContentPath"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content (New-VSCodeSystemInstallWrapper -InstallerFile $installerFileName) `
        -UninstallPs1Content (New-VSCodeSystemUninstallWrapper)

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Visual Studio Code (System) - $version"
        Publisher       = "Microsoft Corporation"
        SoftwareVersion = $version
        DisplayName     = "Visual Studio Code"
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "/VERYSILENT /NORESTART /MERGETASKS=!runcode"
        UninstallCommand = "%ProgramFiles%\Microsoft VS Code\unins000.exe"
        UninstallArgs   = "/VERYSILENT /NORESTART"
        RunningProcess  = @("Code")
        Detection       = @{
            Type          = "File"
            FilePath      = "C:\Program Files\Microsoft VS Code"
            FileName      = "Code.exe"
            PropertyType  = "Version"
            Operator      = "GreaterEquals"
            ExpectedValue = $version
            Is64Bit       = $true
        }
    }

    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"
    return $localContentPath
}

function Invoke-PackageVSCodeSystem {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Visual Studio Code (System) - PACKAGE phase"
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
        $v = Get-LatestVSCodeVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
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
    Write-Log "Visual Studio Code (System) Auto-Packager starting"
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
        Invoke-StageVSCodeSystem
    }
    elseif ($PackageOnly) {
        Invoke-PackageVSCodeSystem
    }
    else {
        Invoke-StageVSCodeSystem
        Invoke-PackageVSCodeSystem
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
