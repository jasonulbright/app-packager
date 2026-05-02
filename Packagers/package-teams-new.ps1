<#
Vendor: Microsoft
App: Microsoft Teams (new client)
CMName: Microsoft Teams
VendorUrl: https://learn.microsoft.com/en-us/microsoftteams/teams-client-bulk-install
CPE: cpe:2.3:a:microsoft:teams:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/en-us/officeupdates/teams-app-versioning
DownloadPageUrl: https://learn.microsoft.com/en-us/microsoftteams/teams-client-bulk-install
UpdateCadenceDays: 14

.SYNOPSIS
    Packages the new Microsoft Teams client for MECM.

.DESCRIPTION
    Downloads the official Teams bootstrapper and x64 offline MSIX, stages
    content to a versioned local folder by reading AppxManifest.xml from the
    MSIX, and creates an MECM Application using script detection against the
    provisioned MSTeams package.

    Supports two-phase operation:
      -StageOnly    Download bootstrapper + MSIX, write wrappers and manifest
      -PackageOnly  Read manifest, copy content to network, create MECM app

    GetLatestVersionOnly reads the public Teams version history and returns
    the newest Windows build that is no longer marked "rolling out". Stage is
    authoritative: it parses the downloaded MSIX version and uses that version
    in the package manifest.

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

$BootstrapperUrl       = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
$TeamsMsixUrl          = "https://go.microsoft.com/fwlink/?linkid=2196106"
$TeamsVersionHistoryUrl = "https://learn.microsoft.com/en-us/officeupdates/teams-app-versioning"

$BootstrapperFileName = "teamsbootstrapper.exe"
$MsixFileName         = "MSTeams-x64.msix"

$VendorFolder = "Microsoft"
$AppFolder    = "Microsoft Teams New Client"
$BaseDownloadRoot = Join-Path $DownloadRoot "TeamsNew"

function Get-LatestTeamsWindowsVersion {
    param([switch]$Quiet)

    Write-Log "Teams version history URL    : $TeamsVersionHistoryUrl" -Quiet:$Quiet

    try {
        $html = (curl.exe -L --fail --silent --show-error $TeamsVersionHistoryUrl) -join ''
        if ($LASTEXITCODE -ne 0) { throw "Failed to query Teams version history." }

        $windowsIndex = $html.IndexOf('<h4 id="windows">Windows</h4>')
        if ($windowsIndex -lt 0) { throw "Could not find public cloud Windows version table." }

        $nextSectionIndex = $html.IndexOf('<h4', $windowsIndex + 1)
        if ($nextSectionIndex -lt 0) { $nextSectionIndex = $html.Length }
        $section = $html.Substring($windowsIndex, $nextSectionIndex - $windowsIndex)

        $rowPattern = '<tr>\s*<td[^>]*>\s*(?<year>\d{4})\s*</td>\s*<td[^>]*>\s*(?<date>.*?)\s*</td>\s*<td[^>]*>\s*(?<version>\d+(?:\.\d+){3})\s*</td>'
        $matches = [regex]::Matches($section, $rowPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
        foreach ($m in $matches) {
            $dateText = ([regex]::Replace($m.Groups['date'].Value, '<[^>]+>', '')).Trim()
            if ($dateText -notmatch 'rolling out') {
                $version = $m.Groups['version'].Value
                Write-Log "Latest Teams Windows version : $version" -Quiet:$Quiet
                return $version
            }
        }

        throw "Could not parse a non-rolling-out Teams Windows version."
    }
    catch {
        Write-Log "Failed to get Teams version: $($_.Exception.Message)" -Level ERROR -Quiet:$Quiet
        return $null
    }
}

function Get-MsixIdentity {
    param([Parameter(Mandatory)][string]$Path)

    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry('AppxManifest.xml')
        if (-not $entry) { throw "AppxManifest.xml not found in MSIX." }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        try { [xml]$manifest = $reader.ReadToEnd() }
        finally { $reader.Dispose() }

        $identity = $manifest.Package.Identity
        if (-not $identity -or [string]::IsNullOrWhiteSpace([string]$identity.Version)) {
            throw "Could not parse MSIX identity version."
        }

        return [pscustomobject]@{
            Name                  = [string]$identity.Name
            Publisher             = [string]$identity.Publisher
            PublisherId           = [string]$identity.PublisherId
            Version               = [string]$identity.Version
            ProcessorArchitecture = [string]$identity.ProcessorArchitecture
        }
    }
    finally {
        $zip.Dispose()
    }
}

function New-TeamsInstallWrapper {
    param(
        [Parameter(Mandatory)][string]$Bootstrapper,
        [Parameter(Mandatory)][string]$Msix
    )

    return (
        ('$bootstrapper = Join-Path $PSScriptRoot ''{0}''' -f $Bootstrapper),
        ('$msixPath = Join-Path $PSScriptRoot ''{0}''' -f $Msix),
        'if (-not (Test-Path -LiteralPath $bootstrapper)) { Write-Error "Missing teamsbootstrapper.exe"; exit 2 }',
        'if (-not (Test-Path -LiteralPath $msixPath)) { Write-Error "Missing MSTeams MSIX"; exit 3 }',
        '$proc = Start-Process -FilePath $bootstrapper -ArgumentList @(''-p'', ''-o'', "`"$msixPath`"") -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
}

function New-TeamsUninstallWrapper {
    param([Parameter(Mandatory)][string]$Bootstrapper)

    return (
        ('$bootstrapper = Join-Path $PSScriptRoot ''{0}''' -f $Bootstrapper),
        'if (Test-Path -LiteralPath $bootstrapper) {',
        '    $proc = Start-Process -FilePath $bootstrapper -ArgumentList @(''-x'', ''-m'') -Wait -PassThru -NoNewWindow',
        '    exit $proc.ExitCode',
        '}',
        'Get-AppxPackage -Name MSTeams -AllUsers | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue',
        'exit 0'
    ) -join "`r`n"
}

function Invoke-StageTeamsNew {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Microsoft Teams (new client) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    if (-not (Test-IsAdmin)) {
        Write-Log "Run PowerShell as Administrator." -Level ERROR
        exit 1
    }

    Initialize-Folder -Path $BaseDownloadRoot

    $localBootstrapper = Join-Path $BaseDownloadRoot $BootstrapperFileName
    $localMsix = Join-Path $BaseDownloadRoot $MsixFileName

    Write-Log "Bootstrapper URL             : $BootstrapperUrl"
    Write-Log "Teams MSIX URL               : $TeamsMsixUrl"
    Write-Log "Local bootstrapper           : $localBootstrapper"
    Write-Log "Local MSIX                   : $localMsix"
    Write-Log ""

    Invoke-DownloadWithRetry -Url $BootstrapperUrl -OutFile $localBootstrapper
    Invoke-DownloadWithRetry -Url $TeamsMsixUrl -OutFile $localMsix

    $identity = Get-MsixIdentity -Path $localMsix
    $version = $identity.Version
    if ($identity.Name -ne 'MSTeams') {
        throw "Downloaded MSIX identity '$($identity.Name)' was not MSTeams."
    }

    Write-Log "MSIX identity name           : $($identity.Name)"
    Write-Log "MSIX version                 : $version"
    Write-Log "MSIX architecture            : $($identity.ProcessorArchitecture)"
    Write-Log ""

    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    Copy-Item -LiteralPath $localBootstrapper -Destination (Join-Path $localContentPath $BootstrapperFileName) -Force -ErrorAction Stop
    Copy-Item -LiteralPath $localMsix -Destination (Join-Path $localContentPath $MsixFileName) -Force -ErrorAction Stop
    Write-Log "Copied bootstrapper + MSIX   : $localContentPath"

    $installPs1 = New-TeamsInstallWrapper -Bootstrapper $BootstrapperFileName -Msix $MsixFileName
    $uninstallPs1 = New-TeamsUninstallWrapper -Bootstrapper $BootstrapperFileName
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installPs1 `
        -UninstallPs1Content $uninstallPs1

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    # Detection targets the provisioned MSIX install path. CCM script-clause
    # detection is invoked under an effective AllSigned policy regardless of the
    # client's LocalMachine policy, so unsigned PS detection scripts fail with
    # 0x87d00327 on lab clients. A versioned file-existence clause sidesteps that
    # entirely; the version is encoded in the AppX package full name and the
    # manifest is regenerated each Stage with the current MSIX version.
    $teamsAppDir = "C:\Program Files\WindowsApps\MSTeams_${version}_x64__8wekyb3d8bbwe"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Microsoft Teams - $version (new client)"
        Publisher       = "Microsoft Corporation"
        SoftwareVersion = $version
        DisplayName     = "Microsoft Teams"
        InstallerFile   = $BootstrapperFileName
        InstallerType   = "EXE"
        InstallArgs     = '-p -o ".\MSTeams-x64.msix"'
        UninstallArgs   = "-x -m"
        RunningProcess  = @("ms-teams", "msteams", "Teams")
        Detection       = @{
            Type         = "File"
            FilePath     = $teamsAppDir
            FileName     = "ms-teams.exe"
            PropertyType = "Existence"
            Is64Bit      = $true
        }
    }

    Set-Content -LiteralPath (Join-Path $BaseDownloadRoot "staged-version.txt") -Value $version -Encoding ASCII -ErrorAction Stop

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"
    return $localContentPath
}

function Invoke-PackageTeamsNew {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "Microsoft Teams (new client) - PACKAGE phase"
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
    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    $manifest = Read-StageManifest -Path $manifestPath

    Write-Log "AppName                      : $($manifest.AppName)"
    Write-Log "Publisher                    : $($manifest.Publisher)"
    Write-Log "SoftwareVersion              : $($manifest.SoftwareVersion)"
    Write-Log "Detection Type               : $($manifest.Detection.Type)"
    Write-Log ""

    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
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

if ($GetLatestVersionOnly) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        $v = Get-LatestTeamsWindowsVersion -Quiet
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
    Write-Log "Microsoft Teams (new client) Auto-Packager starting"
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
        Invoke-StageTeamsNew
    }
    elseif ($PackageOnly) {
        Invoke-PackageTeamsNew
    }
    else {
        Invoke-StageTeamsNew
        Invoke-PackageTeamsNew
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
