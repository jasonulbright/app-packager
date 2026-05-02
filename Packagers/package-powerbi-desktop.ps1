<#
Vendor: Microsoft
App: Power BI Desktop (x64)
CMName: Power BI Desktop
VendorUrl: https://powerbi.microsoft.com/desktop/
CPE: cpe:2.3:a:microsoft:power_bi:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/en-us/power-bi/fundamentals/desktop-latest-update
DownloadPageUrl: https://www.microsoft.com/download/details.aspx?id=58494
UpdateCadenceDays: 30

.SYNOPSIS
    Packages Microsoft Power BI Desktop (x64) for MECM.

.DESCRIPTION
    Reads the official Microsoft Download Center details page for the current
    x64 Power BI Desktop installer, downloads the EXE, stages content to a
    versioned local folder, and creates an MECM Application with file-version
    detection against PBIDesktop.exe.

    Supports two-phase operation:
      -StageOnly    Download EXE, write wrappers and manifest
      -PackageOnly  Read manifest, copy to network, create MECM app

    Microsoft documents the EXE silent switches and deployment properties:
    -quiet, -norestart, ACCEPT_EULA=1, and DISABLE_UPDATE_NOTIFICATION=1.

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (Package phase)
    - Local administrator
    - Write access to FileServerPath (Package phase)
    - Internet access (Stage phase)
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 20,
    [int]$MaximumRuntimeMins = 45,
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

$DownloadCenterUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=58494"

$VendorFolder = "Microsoft"
$AppFolder    = "Power BI Desktop"
$BaseDownloadRoot = Join-Path $DownloadRoot "PowerBIDesktop"

function Get-PowerBIDownloadInfo {
    param([switch]$Quiet)

    Write-Log "Power BI Download Center URL : $DownloadCenterUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $DownloadCenterUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Microsoft Download Center." }

        $jsonMatch = [regex]::Match(
            $html,
            'window\.__DLCDetails__=(?<json>\{.*?\})(?=</script>)',
            [System.Text.RegularExpressions.RegexOptions]::Singleline
        )
        if (-not $jsonMatch.Success) {
            throw "Could not find Download Center metadata block."
        }

        $details = $jsonMatch.Groups['json'].Value | ConvertFrom-Json -ErrorAction Stop
        $file = @($details.dlcDetailsView.downloadFile | Where-Object { $_.name -eq 'PBIDesktopSetup_x64.exe' } | Select-Object -First 1)
        if (-not $file) { throw "Could not find PBIDesktopSetup_x64.exe in Download Center metadata." }

        $version = [string]$file.version
        $url = [string]$file.url
        if ([string]::IsNullOrWhiteSpace($version) -or [string]::IsNullOrWhiteSpace($url)) {
            throw "Power BI metadata was missing version or URL."
        }

        Write-Log "Latest Power BI version      : $version" -Quiet:$Quiet
        return [pscustomobject]@{
            Version     = $version
            FileName    = [string]$file.name
            DownloadUrl = $url
        }
    }
    catch {
        Write-Log "Failed to get Power BI Desktop metadata: $($_.Exception.Message)" -Level ERROR -Quiet:$Quiet
        return $null
    }
}

function New-PowerBIInstallWrapper {
    param([Parameter(Mandatory)][string]$InstallerFile)

    return (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFile),
        'if (-not (Test-Path -LiteralPath $exePath)) { Write-Error "Missing Power BI Desktop installer"; exit 2 }',
        '$args = @(''-quiet'', ''-norestart'', ''ACCEPT_EULA=1'', ''DISABLE_UPDATE_NOTIFICATION=1'')',
        '$proc = Start-Process -FilePath $exePath -ArgumentList $args -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}

function New-PowerBIUninstallWrapper {
    param([Parameter(Mandatory)][string]$InstallerFile)

    return (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFile),
        'if (-not (Test-Path -LiteralPath $exePath)) { exit 0 }',
        '$args = @(''-uninstall'', ''-quiet'', ''-norestart'')',
        '$proc = Start-Process -FilePath $exePath -ArgumentList $args -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}

function Invoke-StagePowerBIDesktop {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Power BI Desktop (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $info = Get-PowerBIDownloadInfo
    if (-not $info) { throw "Could not resolve Power BI Desktop download info." }

    $version = $info.Version
    $installerFileName = $info.FileName
    $downloadUrl = $info.DownloadUrl
    $localInstaller = Join-Path $BaseDownloadRoot $installerFileName

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $downloadUrl"
    Write-Log "Local installer              : $localInstaller"
    Write-Log ""

    if (-not (Test-Path -LiteralPath $localInstaller)) {
        Invoke-DownloadWithRetry -Url $downloadUrl -OutFile $localInstaller
    }
    else {
        Write-Log "Local installer exists. Skipping download."
    }

    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    Copy-Item -LiteralPath $localInstaller -Destination (Join-Path $localContentPath $installerFileName) -Force -ErrorAction Stop
    Write-Log "Copied installer to stage    : $localContentPath"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content (New-PowerBIInstallWrapper -InstallerFile $installerFileName) `
        -UninstallPs1Content (New-PowerBIUninstallWrapper -InstallerFile $installerFileName) `
        -InstallBatExitCode '3010'

    $detectionPath = "{0}\Microsoft Power BI Desktop\bin" -f $env:ProgramFiles
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Power BI Desktop - $version"
        Publisher       = "Microsoft Corporation"
        SoftwareVersion = $version
        DisplayName     = "Power BI Desktop"
        InstallerFile   = $installerFileName
        InstallerType   = "EXE"
        InstallArgs     = "-quiet -norestart ACCEPT_EULA=1 DISABLE_UPDATE_NOTIFICATION=1"
        UninstallArgs   = "-uninstall -quiet -norestart"
        RunningProcess  = @("PBIDesktop")
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "PBIDesktop.exe"
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

function Invoke-PackagePowerBIDesktop {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Power BI Desktop (x64) - PACKAGE phase"
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
        $info = Get-PowerBIDownloadInfo -Quiet
        if (-not $info) { exit 1 }
        Write-Output $info.Version
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
    Write-Log "Power BI Desktop (x64) Auto-Packager starting"
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
        Invoke-StagePowerBIDesktop
    }
    elseif ($PackageOnly) {
        Invoke-PackagePowerBIDesktop
    }
    else {
        Invoke-StagePowerBIDesktop
        Invoke-PackagePowerBIDesktop
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
