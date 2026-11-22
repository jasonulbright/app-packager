<#
Vendor: TeamViewer
App: TeamViewer Host (x64)
CMName: TeamViewer Host
VendorUrl: https://www.teamviewer.com/
CPE: cpe:2.3:a:teamviewer:teamviewer:*:*:*:*:*:*:*:*
ReleaseNotesUrl: https://www.teamviewer.com/en-us/whats-new/
DownloadPageUrl: https://www.teamviewer.com/en-us/download/windows/
SecurityAdvisoryUrl: https://www.teamviewer.com/en/trust-center/security-bulletins/
IconSource: Installer
LocalSource: Required
SupportsInstallerTypes: EXE, MSI
WsusSupport: Yes
ArpName: TeamViewer Host

.SYNOPSIS
    Packages TeamViewer Host (x64) for ConfigMgr, as an MSI or as the vendor EXE.

.DESCRIPTION
    The installer type is a setting of the application (Options > Applications
    > All Applications > Installer Type). MSI is the default and is described
    below. EXE downloads TeamViewer_Host_Setup_x64.exe, reads the version from
    its file properties, installs with /S, uninstalls with the NSIS uninstall.exe
    /S and detects TeamViewer.exe by file version. TeamViewer documents the
    silent install for the MSI (msiexec /qn plus properties) and does not
    document an EXE switch, so the EXE type does not pass the properties of
    teamviewer-host-config.json; the packager logs a warning for each one.

    Stages the TeamViewer Host MSI from the installer source folder of this
    packager (Options > Applications > Local Installer Sources), creates the
    content wrappers and a ConfigMgr Application with detection by the
    product's Add/Remove Programs entry.

    TeamViewer gives the Host MSI only to signed-in Management Console users
    (Design and Deploy), so the packager has no download. Put the MSI in the
    folder: the file TeamViewer_Host.msi from the Management Console download
    (unzip the download first when it is a .zip). When the folder holds more
    than one match, the file with the highest ProductVersion is staged. The
    version comes from that file, not from a web page.

    Supports two-phase operation:
      -StageOnly    Read the MSI from the folder, derive the version, write manifest
      -PackageOnly  Read manifest, copy to network, create ConfigMgr application

    TeamViewer Host is the unattended-access variant of TeamViewer. It is
    deployed to endpoints that need to be remotely managed by IT without a
    user initiating the session. The Full client (package-teamviewer.ps1) is
    deployed to IT staff who initiate support sessions.

    The values of teamviewer-host-config.json become MSI properties on the
    install command line: ApiToken (APITOKEN), CustomConfigId (CUSTOMCONFIGID),
    AssignmentOptions (ASSIGNMENTOPTIONS) and RemoveDesktopShortcut
    (DESKTOPSHORTCUTS=0). An empty
    value is not passed. The values are written to the install script of the
    package in clear text.

    GetLatestVersionOnly has no web source to read. It prints the ProductVersion
    of the newest matching MSI in the installer source folder, and it fails with
    a message that names the folder when the folder holds no matching file.

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (for example, "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the ConfigMgr application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\TeamViewer\TeamViewer Host\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging installers.
    Each packager creates a subfolder under this path (for example, <DownloadRoot>\TeamViewerHost).
    Default: C:\temp\ap

.PARAMETER SourceFolder
    The folder that holds the TeamViewer Host MSI. It replaces the folder saved in
    Options for this run.

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the ConfigMgr deployment type.
    Default: 15

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the ConfigMgr deployment type.
    Default: 30

.PARAMETER StageOnly
    Runs only the Stage phase: read the MSI from the installer source folder,
    derive the version and the detection from it, generate content wrappers and
    the stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create ConfigMgr application with Add/Remove Programs detection.

.PARAMETER GetLatestVersionOnly
    Outputs the ProductVersion of the newest matching MSI in the installer
    source folder and exits. No ConfigMgr changes are made.

.REQUIREMENTS
    - PowerShell 5.1
    - Configuration Manager console installed (ConfigurationManager PowerShell module available)
    - RBAC permissions to create Applications and Deployment Types
    - Write access to FileServerPath
    - The TeamViewer Host MSI in the installer source folder
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
    [string]$SourceFolder = "",
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
# The installer comes from the installer source folder (no download). The
# name rule: the Management Console file TeamViewer_Host.msi, with a suffix
# that a browser adds to a repeated download, for example "TeamViewer_Host (1).msi".
$PackagerName      = "package-teamviewerhost"
$InstallerFilter   = "TeamViewer_Host*.msi"
$InstallerFileName = "TeamViewer_Host.msi"
$FolderHint = "Put the TeamViewer Host MSI (TeamViewer_Host.msi) in that folder. TeamViewer gives it only in the Management Console, under Design and Deploy; unzip the download first when it is a .zip."

# The EXE installer type downloads the vendor NSIS installer. The MSI type reads the MSI above from the installer source folder.
$ExeDownloadUrl = "https://dl.teamviewer.com/download/TeamViewer_Host_Setup_x64.exe"
$ExeFileName    = "TeamViewer_Host_Setup_x64.exe"

$VendorFolder = "TeamViewer"
$AppFolder    = "TeamViewer Host"

$BaseDownloadRoot = Join-Path $DownloadRoot "TeamViewerHost"
$ConfigFile       = Resolve-PackagerRunFile -Name "teamviewer-host-config.json" -Default (Join-Path $PSScriptRoot "teamviewer-host-config.json")

# --- Functions ---


function Resolve-TeamViewerHostMsi {
    <#
    .SYNOPSIS
        The MSI in the installer source folder with the highest ProductVersion.
        Throws a message that names the folder and the file to put there.
    #>
    return (Resolve-LocalSourceInstaller -PackagerName $PackagerName -Filter $InstallerFilter -Override $SourceFolder -Hint $FolderHint)
}


function Read-TeamViewerHostConfig {
    <#
    .SYNOPSIS
        Reads teamviewer-host-config.json. Returns $null when the file is absent,
        empty or unreadable (an unreadable file logs a warning).
    #>
    if (-not (Test-Path -LiteralPath $ConfigFile)) { return $null }
    try {
        $raw = Get-Content -LiteralPath $ConfigFile -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return ($raw | ConvertFrom-Json -ErrorAction Stop)
    }
    catch {
        Write-Log "teamviewer-host-config.json unreadable: $($_.Exception.Message)" -Level WARN
        return $null
    }
}


function Get-TeamViewerHostMsiArguments {
    <#
    .SYNOPSIS
        The MSI properties that the values of teamviewer-host-config.json add to
        the msiexec command line. One element per argument. An empty value adds
        nothing.

    .DESCRIPTION
        ApiToken -> APITOKEN, CustomConfigId -> CUSTOMCONFIGID,
        AssignmentOptions -> ASSIGNMENTOPTIONS (quoted: the value holds spaces),
        RemoveDesktopShortcut -> DESKTOPSHORTCUTS=0. A value with a double quote or a line break is
        refused, because it would end the quoted property early.
    #>
    param([AllowNull()]$Config)

    $arguments = New-Object System.Collections.Generic.List[string]
    if (-not $Config) { return @() }

    $text = {
        param([string]$Name)
        $property = $Config.PSObject.Properties[$Name]
        if (-not $property -or $null -eq $property.Value) { return '' }
        $value = ([string]$property.Value).Trim()
        if ($value -match '["\r\n]') { throw "teamviewer-host-config.json: the value of $Name holds a double quote or a line break." }
        return $value
    }

    $token = & $text 'ApiToken'
    if ($token) {
        if ($token -notmatch '^\S+$') { throw 'teamviewer-host-config.json: the value of ApiToken holds a space.' }
        $arguments.Add("APITOKEN=$token")
    }
    $configId = & $text 'CustomConfigId'
    if ($configId) {
        if ($configId -notmatch '^\S+$') { throw 'teamviewer-host-config.json: the value of CustomConfigId holds a space.' }
        $arguments.Add("CUSTOMCONFIGID=$configId")
    }
    $options = & $text 'AssignmentOptions'
    if ($options) {
        if ($options.EndsWith('\')) { throw 'teamviewer-host-config.json: the value of AssignmentOptions ends with a backslash.' }
        $arguments.Add("ASSIGNMENTOPTIONS=`"$options`"")
    }
    if ($Config.PSObject.Properties['RemoveDesktopShortcut'] -and ([string]$Config.RemoveDesktopShortcut) -match '^(true|1)$') {
        $arguments.Add('DESKTOPSHORTCUTS=0')
    }
    return $arguments.ToArray()
}


function Get-RequestedTeamViewerHostInstallerType {
    <#
    .SYNOPSIS
        The installer type of this run: the application setting (EXE or MSI), MSI when the setting is the packager default.
    #>
    $type = Get-RequestedInstallerType
    if (-not $type) { return 'MSI' }
    return $type
}


function Get-TeamViewerHostExeVersion {
    <#
    .SYNOPSIS
        Reads the product version from a TeamViewer Host EXE as a dotted version string.

    .DESCRIPTION
        The NSIS installer leaves the version resource string table empty, so FileVersionInfo.ProductVersion and FileVersion come
        back blank. The fixed block of the same resource is populated, and its four Product parts match the vendor MSI ProductVersion.
        The string table is tried first, then the fixed block. Both empty throws.
    #>
    param([Parameter(Mandatory)][string]$ExePath)

    $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($ExePath)
    $v  = $vi.ProductVersion
    if ([string]::IsNullOrWhiteSpace($v)) { $v = $vi.FileVersion }
    if (-not [string]::IsNullOrWhiteSpace($v)) { return $v.Trim() }

    $fixed = "{0}.{1}.{2}.{3}" -f $vi.ProductMajorPart, $vi.ProductMinorPart, $vi.ProductBuildPart, $vi.ProductPrivatePart
    if ($fixed -ne '0.0.0.0') { return $fixed }

    throw "Could not read a product version from $ExePath - the version resource string table and its fixed block are both empty."
}


function Write-IgnoredConfigWarnings {
    <#
    .SYNOPSIS
        Logs a warning for each teamviewer-host-config.json key that is an MSI property. The EXE type does not pass them.
    #>
    $cfg = Read-TeamViewerHostConfig
    if (-not $cfg) { return }

    $ignored = @()
    if (-not [string]::IsNullOrWhiteSpace([string]$cfg.ApiToken))          { $ignored += 'ApiToken (APITOKEN)' }
    if (-not [string]::IsNullOrWhiteSpace([string]$cfg.CustomConfigId))    { $ignored += 'CustomConfigId (CUSTOMCONFIGID)' }
    if (-not [string]::IsNullOrWhiteSpace([string]$cfg.AssignmentOptions)) { $ignored += 'AssignmentOptions (ASSIGNMENTOPTIONS)' }
    if ($cfg.PSObject.Properties['RemoveDesktopShortcut'] -and ([string]$cfg.RemoveDesktopShortcut) -match '^(true|1)$') { $ignored += 'RemoveDesktopShortcut (DESKTOPSHORTCUTS)' }

    if ($ignored.Count -gt 0) {
        Write-Log "teamviewer-host-config.json holds MSI properties that the EXE installer type does not pass:" -Level WARN
        foreach ($k in $ignored) { Write-Log "  - $k" -Level WARN }
        Write-Log "These values are ignored for the EXE type. Choose the MSI installer type to use them." -Level WARN
    }
}


function Get-LatestTeamViewerHostVersion {
    param([switch]$Quiet)

    if ((Get-RequestedTeamViewerHostInstallerType) -eq 'EXE') {
        Write-Log "TeamViewer Host EXE URL      : $ExeDownloadUrl" -Quiet:$Quiet
        try {
            Initialize-Folder -Path $BaseDownloadRoot
            $localExe = Join-Path $BaseDownloadRoot $ExeFileName
            Invoke-CachedDownload -Url $ExeDownloadUrl -OutFile $localExe -Quiet:$Quiet
            $version = Get-TeamViewerHostExeVersion -ExePath $localExe
            Write-Log "Latest TeamViewer Host ver   : $version" -Quiet:$Quiet
            return $version
        }
        catch {
            Write-Log "Failed to get TeamViewer Host version: $($_.Exception.Message)" -Level ERROR
            return $null
        }
    }

    try {
        $localMsi = Resolve-TeamViewerHostMsi
        Write-Log "Installer source             : $localMsi" -Quiet:$Quiet
        $version = (Get-MsiPropertyMap -MsiPath $localMsi)['ProductVersion']
        if ([string]::IsNullOrWhiteSpace($version)) { throw "MSI ProductVersion missing: $localMsi" }
        Write-Log "Latest TeamViewer Host ver   : $version" -Quiet:$Quiet
        return $version
    }
    catch {
        # TeamViewer publishes no version a script can read; say where the version comes from.
        Write-Log "TeamViewer Host has no web version source: the version is read from the MSI in the installer source folder. $($_.Exception.Message)" -Level ERROR
        return $null
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageTeamViewerHostExe {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeamViewer Host (x64, EXE) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    $localExe = Join-Path $BaseDownloadRoot $ExeFileName
    Write-Log "Local EXE path               : $localExe"
    Invoke-CachedDownload -Url $ExeDownloadUrl -OutFile $localExe

    $version = Get-TeamViewerHostExeVersion -ExePath $localExe
    Write-Log "EXE ProductVersion           : $version"
    Write-Log ""

    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath
    Copy-Item -LiteralPath $localExe -Destination (Join-Path $localContentPath $ExeFileName) -Force -ErrorAction Stop
    Write-Log "Copied EXE to staged folder  : $localContentPath"

    Write-IgnoredConfigWarnings

    # Install: stop a running TeamViewer first (the installer cannot replace an open exe), run the NSIS installer with /S, then close
    # a GUI that the post-install step may launch in the SYSTEM session. Uninstall: the NSIS uninstaller at the Host install path.
    $installContent = (
        'Stop-Process -Name "TeamViewer","tv_w32","tv_x64" -Force -ErrorAction SilentlyContinue',
        ('$exePath = Join-Path $PSScriptRoot ''{0}''' -f $ExeFileName),
        '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'Stop-Process -Name "TeamViewer","tv_w32","tv_x64" -Force -ErrorAction SilentlyContinue',
        'exit $proc.ExitCode'
    ) -join "`r`n"
    $uninstallContent = (
        'Stop-Process -Name "TeamViewer","tv_w32","tv_x64" -Force -ErrorAction SilentlyContinue',
        '$uninstall = Join-Path $env:ProgramFiles ''TeamViewer\uninstall.exe''',
        'if (-not (Test-Path -LiteralPath $uninstall)) { exit 0 }',
        '$proc = Start-Process -FilePath $uninstall -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
        'exit $proc.ExitCode'
    ) -join "`r`n"
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $installContent `
        -UninstallPs1Content $uninstallContent

    $detectionPath = "{0}\TeamViewer" -f $env:ProgramFiles
    Write-Log "Detection                    : $detectionPath\TeamViewer.exe version >= $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "TeamViewer Host $version"
        Publisher       = "TeamViewer"
        SoftwareVersion = $version
        InstallerFile   = $ExeFileName
        InstallerType   = "EXE"
        InstallArgs     = "/S"
        UninstallArgs   = "/S"
        RunningProcess  = @("TeamViewer_Service", "TeamViewer")
        Detection       = @{
            Type          = "File"
            FilePath      = $detectionPath
            FileName      = "TeamViewer.exe"
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


function Invoke-StageTeamViewerHost {
    if ((Get-RequestedTeamViewerHostInstallerType) -eq 'EXE') { return (Invoke-StageTeamViewerHostExe) }
    return (Invoke-StageTeamViewerHostMsi)
}


function Invoke-StageTeamViewerHostMsi {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeamViewer Host (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    Initialize-Folder -Path $BaseDownloadRoot

    # --- Installer from the installer source folder ---
    $localMsi = Resolve-TeamViewerHostMsi
    Write-Log "Installer source             : $localMsi"

    $signature = Get-AuthenticodeSignature -LiteralPath $localMsi
    if ($signature.Status -ne 'Valid') {
        throw "Installer signature is $($signature.Status), not Valid: $localMsi"
    }
    Write-Log "Installer signer             : $($signature.SignerCertificate.Subject)"

    # --- Version and product code from the MSI ---
    $props = Get-MsiPropertyMap -MsiPath $localMsi
    $productName = $props["ProductName"]
    $version     = $props["ProductVersion"]
    $productCode = $props["ProductCode"]
    if ([string]::IsNullOrWhiteSpace($version))     { throw "MSI ProductVersion missing: $localMsi" }
    if ([string]::IsNullOrWhiteSpace($productCode)) { throw "MSI ProductCode missing: $localMsi" }

    Write-Log "MSI ProductName              : $productName"
    Write-Log "MSI ProductVersion           : $version"
    Write-Log "MSI Manufacturer             : $($props['Manufacturer'])"
    Write-Log "MSI ProductCode              : $productCode"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedMsi = Join-Path $localContentPath $InstallerFileName
    Copy-Item -LiteralPath $localMsi -Destination $stagedMsi -Force -ErrorAction Stop
    Write-Log "Copied MSI to staged folder  : $stagedMsi"

    # --- MSI properties from teamviewer-host-config.json ---
    $extraArguments = @(Get-TeamViewerHostMsiArguments -Config (Read-TeamViewerHostConfig))
    if ($extraArguments.Count -gt 0) {
        # The names only: the token and the ids stay out of the log.
        $names = @($extraArguments | ForEach-Object { ($_ -split '=', 2)[0] })
        Write-Log "MSI properties from config   : $($names -join ', ')"
    }

    # --- Generate content wrappers ---
    $wrapperContent = New-MsiWrapperContent -MsiFileName $InstallerFileName -ExtraInstallArgs $extraArguments

    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    # Detection reads the product's Add/Remove Programs entry in both registry
    # views, so it does not depend on the bitness of the package.
    $arpEntry = @{
        Type              = "ArpEntry"
        View              = "Both"
        DisplayNamePrefix = "TeamViewer Host"
        WindowsInstaller  = $true
        Version           = $version
    }
    Write-Log "Detection                    : Add/Remove Programs entry 'TeamViewer Host*' (Windows Installer, both views), DisplayVersion >= $version"
    Write-Log ""

    # The WSUS publish builds its command line from this field, so it holds real values only. The token stays out of
    # it (it is in install.ps1 only); a WSUS package installs without the token.
    $installArgsText = "/qn /norestart"
    $manifestArguments = @($extraArguments | Where-Object { $_ -notlike 'APITOKEN=*' })
    if ($manifestArguments.Count -gt 0) { $installArgsText += ' ' + ($manifestArguments -join ' ') }

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "TeamViewer Host $version"
        Publisher       = "TeamViewer"
        SoftwareVersion = $version
        InstallerFile   = $InstallerFileName
        InstallerType   = "MSI"
        InstallArgs     = $installArgsText
        UninstallArgs   = "/qn /norestart"
        ProductCode     = $productCode
        RunningProcess  = @("TeamViewer_Service", "TeamViewer")
        Detection       = @{
            Type           = "Script"
            ScriptLanguage = "PowerShell"
            ScriptText     = (New-ArpEntryDetectionScript -Entry $arpEntry)
        }
        WsusDetection   = $arpEntry
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

function Invoke-PackageTeamViewerHost {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "TeamViewer Host (x64) - PACKAGE phase"
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

    # --- ConfigMgr application ---
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
        $v = Get-LatestTeamViewerHostVersion -Quiet
        if (-not $v) { exit 1 }
        Write-Output $v
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
    Write-Log "TeamViewer Host (x64) Auto-Packager starting"
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
        Invoke-StageTeamViewerHost
    }
    elseif ($PackageOnly) {
        Invoke-PackageTeamViewerHost
    }
    else {
        Invoke-StageTeamViewerHost
        Invoke-PackageTeamViewerHost
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-teamviewerhost'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
