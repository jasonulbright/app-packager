<#
Vendor: Microsoft
App: SQL Server Management Studio 22
CMName: SQL Server Management Studio
VendorUrl: https://learn.microsoft.com/en-us/ssms/sql-server-management-studio-ssms
CPE: cpe:2.3:a:microsoft:sql_server_management_studio:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/en-us/ssms/release-history
DownloadPageUrl: https://learn.microsoft.com/en-us/ssms/download-sql-server-management-studio-ssms
UpdateCadenceDays: 14

.SYNOPSIS
    Packages SQL Server Management Studio 22 for MECM.

.DESCRIPTION
    Downloads the official SSMS 22 release-channel bootstrapper, stages it to
    a versioned local folder, and creates an MECM Application with file-version
    detection against Ssms.exe.

    SSMS 22 uses the Visual Studio Installer bootstrapper model. Silent switch
    choices are read from Packagers\packager-preferences.json under
    SSMSInstallOptions, which is written by the Packager Preferences UI.

    Supports two-phase operation:
      -StageOnly    Download bootstrapper, write wrappers and manifest
      -PackageOnly  Read manifest, copy to network, create MECM app

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (Package phase)
    - Local administrator
    - Write access to FileServerPath (Package phase)
    - Internet access during install unless using a layout in a future packager
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [int]$EstimatedRuntimeMins = 30,
    [int]$MaximumRuntimeMins = 90,
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

$CommandLineDocUrl = "https://learn.microsoft.com/en-us/ssms/install/command-line-parameters"
$BootstrapperUrl   = "https://aka.ms/ssms/22/release/vs_SSMS.exe"
$BootstrapperFile  = "vs_SSMS.exe"

$VendorFolder = "Microsoft"
$AppFolder    = "SQL Server Management Studio 22"
$BaseDownloadRoot = Join-Path $DownloadRoot "SSMS22"

$DefaultInstallPath = "C:\Program Files\Microsoft SQL Server Management Studio 22\Release"

function Get-LatestSsmsRelease {
    param([switch]$Quiet)

    Write-Log "SSMS command-line doc URL    : $CommandLineDocUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $CommandLineDocUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query SSMS command-line documentation." }

        $m = [regex]::Match(
            $html,
            '<td>\s*Release\s*</td>\s*<td>\s*(?<version>\d+(?:\.\d+){2})\s*</td>\s*<td>\s*<a[^>]+href="(?<url>[^"]+)"[^>]*>\s*SQL Server Management Studio\s*</a>',
            [System.Text.RegularExpressions.RegexOptions]::Singleline
        )
        if (-not $m.Success) {
            $m = [regex]::Match($html, 'Release\s+(?<version>\d+(?:\.\d+){2})', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        }
        if (-not $m.Success) { throw "Could not parse SSMS release version." }

        $version = $m.Groups['version'].Value
        $url = $BootstrapperUrl
        if ($m.Groups['url'] -and -not [string]::IsNullOrWhiteSpace($m.Groups['url'].Value)) {
            $candidate = [string]$m.Groups['url'].Value
            if ($candidate -match '^https?://') { $url = $candidate }
        }

        Write-Log "Latest SSMS version          : $version" -Quiet:$Quiet
        return [pscustomobject]@{
            Version     = $version
            FileName    = $BootstrapperFile
            DownloadUrl = $url
        }
    }
    catch {
        Write-Log "Failed to get SSMS version: $($_.Exception.Message)" -Level ERROR -Quiet:$Quiet
        return $null
    }
}

function Get-SsmsInstallOptions {
    $defaults = [ordered]@{
        UIMode              = "Quiet"
        DownloadThenInstall = $true
        NoUpdateInstaller   = $false
        IncludeRecommended  = $false
        IncludeOptional     = $false
        RemoveOos           = $true
        ForceClose          = $false
        InstallPath         = ""
    }

    try {
        $prefs = Get-PackagerPreferences
        if ($prefs -and $prefs.SSMSInstallOptions) {
            $cfg = $prefs.SSMSInstallOptions
            if ($cfg.UIMode -in @('Quiet','Passive')) { $defaults.UIMode = [string]$cfg.UIMode }
            foreach ($prop in @('DownloadThenInstall','NoUpdateInstaller','IncludeRecommended','IncludeOptional','RemoveOos','ForceClose')) {
                if ($null -ne $cfg.$prop) { $defaults[$prop] = [bool]$cfg.$prop }
            }
            if ($null -ne $cfg.InstallPath) { $defaults.InstallPath = [string]$cfg.InstallPath }
        }
    }
    catch {
        Write-Log "Could not read SSMSInstallOptions; using defaults: $($_.Exception.Message)" -Level WARN
    }

    return [pscustomobject]$defaults
}

function Get-SsmsInstallPath {
    param([Parameter(Mandatory)][pscustomobject]$Options)

    if (-not [string]::IsNullOrWhiteSpace([string]$Options.InstallPath)) {
        return ([string]$Options.InstallPath).Trim()
    }
    return $DefaultInstallPath
}

function Get-SsmsInstallArgs {
    param([Parameter(Mandatory)][pscustomobject]$Options)

    $args = @()
    if ($Options.UIMode -eq 'Passive') { $args += '--passive' } else { $args += '--quiet' }
    $args += '--norestart'
    if ($Options.DownloadThenInstall -eq $true) { $args += '--downloadThenInstall' }
    if ($Options.NoUpdateInstaller -eq $true)  { $args += '--noUpdateInstaller' }
    if ($Options.IncludeRecommended -eq $true) { $args += '--includeRecommended' }
    if ($Options.IncludeOptional -eq $true)    { $args += '--includeOptional' }
    if ($Options.RemoveOos -eq $true)          { $args += @('--removeOos', 'true') }
    if ($Options.ForceClose -eq $true)         { $args += '--force' }

    $installPath = ([string]$Options.InstallPath).Trim()
    if (-not [string]::IsNullOrWhiteSpace($installPath)) {
        $args += @('--installPath', $installPath)
    }

    return [string[]]$args
}

function Get-SsmsUninstallArgs {
    param([Parameter(Mandatory)][pscustomobject]$Options)

    $args = @('uninstall', '--installPath', (Get-SsmsInstallPath -Options $Options))
    if ($Options.UIMode -eq 'Passive') { $args += '--passive' } else { $args += '--quiet' }
    $args += '--norestart'
    if ($Options.ForceClose -eq $true) { $args += '--force' }
    return [string[]]$args
}

function ConvertTo-CommandLinePreview {
    param([Parameter(Mandatory)][string[]]$Arguments)

    return ($Arguments | ForEach-Object {
        if ($_ -match '[\s"]') { '"' + ($_ -replace '"','\"') + '"' } else { $_ }
    }) -join ' '
}

function New-SsmsInstallWrapper {
    param(
        [Parameter(Mandatory)][string]$InstallerFile,
        [Parameter(Mandatory)][string[]]$Arguments
    )

    $argLiteral = ($Arguments | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ', '
    return (
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFile),
        'if (-not (Test-Path -LiteralPath $exePath)) { Write-Error "Missing SSMS bootstrapper"; exit 2 }',
        ('$args = @({0})' -f $argLiteral),
        '$proc = Start-Process -FilePath $exePath -ArgumentList $args -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}

function New-SsmsUninstallWrapper {
    param(
        [Parameter(Mandatory)][string]$InstallerFile,
        [Parameter(Mandatory)][string[]]$Arguments
    )

    $argLiteral = ($Arguments | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ', '
    return (
        '$setupExe = Join-Path ${env:ProgramFiles(x86)} ''Microsoft Visual Studio\Installer\setup.exe''',
        ('$fallbackBootstrapper = Join-Path $PSScriptRoot ''{0}''' -f $InstallerFile),
        '$exePath = if (Test-Path -LiteralPath $setupExe) { $setupExe } else { $fallbackBootstrapper }',
        'if (-not (Test-Path -LiteralPath $exePath)) { exit 0 }',
        ('$args = @({0})' -f $argLiteral),
        '$proc = Start-Process -FilePath $exePath -ArgumentList $args -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}

function Invoke-StageSsms {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SQL Server Management Studio 22 - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $release = Get-LatestSsmsRelease
    if (-not $release) { throw "Could not resolve SSMS release info." }

    $options = Get-SsmsInstallOptions
    $installArgs = Get-SsmsInstallArgs -Options $options
    $uninstallArgs = Get-SsmsUninstallArgs -Options $options
    $installPath = Get-SsmsInstallPath -Options $options

    $version = $release.Version
    $localBootstrapper = Join-Path $BaseDownloadRoot $BootstrapperFile

    Write-Log "Version                      : $version"
    Write-Log "Download URL                 : $($release.DownloadUrl)"
    Write-Log "Local bootstrapper           : $localBootstrapper"
    Write-Log "Install path                 : $installPath"
    Write-Log "Install arguments            : $(ConvertTo-CommandLinePreview -Arguments $installArgs)"
    Write-Log ""

    Invoke-DownloadWithRetry -Url $release.DownloadUrl -OutFile $localBootstrapper

    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    Copy-Item -LiteralPath $localBootstrapper -Destination (Join-Path $localContentPath $BootstrapperFile) -Force -ErrorAction Stop
    Write-Log "Copied bootstrapper to stage : $localContentPath"

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content (New-SsmsInstallWrapper -InstallerFile $BootstrapperFile -Arguments $installArgs) `
        -UninstallPs1Content (New-SsmsUninstallWrapper -InstallerFile $BootstrapperFile -Arguments $uninstallArgs)

    $detectionPath = Join-Path $installPath "Common7\IDE"
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "SQL Server Management Studio 22 - $version"
        Publisher       = "Microsoft Corporation"
        SoftwareVersion = $version
        DisplayName     = "SQL Server Management Studio"
        InstallerFile   = $BootstrapperFile
        InstallerType   = "EXE"
        InstallArgs     = (ConvertTo-CommandLinePreview -Arguments $installArgs)
        UninstallArgs   = (ConvertTo-CommandLinePreview -Arguments $uninstallArgs)
        RunningProcess  = @("Ssms")
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "Ssms.exe"
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

function Invoke-PackageSsms {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SQL Server Management Studio 22 - PACKAGE phase"
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
        $release = Get-LatestSsmsRelease -Quiet
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
    Write-Log "SQL Server Management Studio 22 Auto-Packager starting"
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
        Invoke-StageSsms
    }
    elseif ($PackageOnly) {
        Invoke-PackageSsms
    }
    else {
        Invoke-StageSsms
        Invoke-PackageSsms
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
