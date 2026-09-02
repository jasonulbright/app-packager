<#
Vendor: Microsoft
App: SQL Server 2022 Express (x64)
CMName: Microsoft SQL Server 2022 Express
VendorUrl: https://www.microsoft.com/sql-server/sql-server-downloads
CPE: cpe:2.3:a:microsoft:sql_server:2022:*:*:*:express:*:*:*
ReleaseNotesUrl: https://learn.microsoft.com/sql/database-engine/configure-windows/sql-server-version-updates
DownloadPageUrl: https://www.microsoft.com/download/details.aspx?id=104781
UpdateCadenceDays: 180

.SYNOPSIS
    Packages Microsoft SQL Server 2022 Express (x64) for MECM.

.DESCRIPTION
    Downloads the SQLEXPR_x64_ENU.exe Express Core media box, stages content to
    a versioned local folder, and creates an MECM Application with
    registry-based detection on the instance Setup key.

    Two installers exist for Express. SQL2022-SSEI-Expr.exe is a small
    downloader that does not forward setup parameters to the media it fetches;
    only the media box accepts them. This packager therefore ships the media box
    itself, so the deployment is a single self-contained payload with no
    client-side download and no elevation needed at stage time. To refresh the
    media by hand when the Download Center URL changes, run
    SQL2022-SSEI-Expr.exe /ACTION=Download /MEDIAPATH=<folder> /MEDIATYPE=Core
    /QUIET from an elevated prompt and point MediaUrl at the result.

    The generated install wrapper extracts the media box and runs setup.exe
    with the full parameter set: Database Engine only, instance SQLEXPRESS,
    Windows authentication, BUILTIN\Administrators as sysadmin.

    Supports two-phase operation:
      -StageOnly    Download, generate content wrappers, write manifest
      -PackageOnly  Read manifest, copy to network, create MECM application

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").
    The PSDrive is assumed to already exist in the session.

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\SQL Server 2022 Express\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging downloaded installers.
    Each packager creates a subfolder under this path (e.g., <DownloadRoot>\SqlServer2022Express).
    Default: C:\temp\ap

.PARAMETER InstanceName
    Database Engine instance name written into the install wrapper and the
    detection key. Default: SQLEXPRESS

.PARAMETER SysAdminAccounts
    Principal added to the sysadmin server role at install time.
    Default: BUILTIN\Administrators

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for the MECM deployment type.
    Default: 30

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for the MECM deployment type.
    Default: 90

.PARAMETER StageOnly
    Runs only the Stage phase: download media, generate content wrappers and
    stage manifest.

.PARAMETER PackageOnly
    Runs only the Package phase: read stage manifest, copy content to network,
    create MECM application with registry-based detection.

.PARAMETER GetLatestVersionOnly
    Outputs only the available SQL Server 2022 Express media version and exits.

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
    [ValidatePattern('^[A-Za-z][A-Za-z0-9_$]{0,15}$')]
    [string]$InstanceName = "SQLEXPRESS",
    [string]$SysAdminAccounts = "BUILTIN\Administrators",
    [int]$EstimatedRuntimeMins = 30,
    [int]$MaximumRuntimeMins = 90,
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
# Express Core media for the SQL Server 2022 Download Center package. The
# fwlink redirector that once served 2022 now serves the 2025 downloader, so
# the version-bearing Download Center path is used instead.
$MediaUrl      = "https://download.microsoft.com/download/3/8/d/38de7036-2433-4207-8eae-06e247e17b25/SQLEXPR_x64_ENU.exe"
$MediaFileName = "SQLEXPR_x64_ENU.exe"

# Setup writes the instance Setup key under MSSQL16.<instance> for the 2022
# (16.x) release family.
$InstanceKeyPrefix = "MSSQL16"

$VendorFolder = "Microsoft"
$AppFolder    = "SQL Server 2022 Express"

$BaseDownloadRoot = Join-Path $DownloadRoot "SqlServer2022Express"

# --- Functions ---


function Assert-PayloadIsExecutable {
    <#
    .SYNOPSIS
        Throws unless the downloaded file starts with the PE 'MZ' signature.
    .DESCRIPTION
        The Download Center answers 200 with an HTML page for package paths
        that have been retired, which would otherwise stage as a valid-looking
        media box.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $bytes = Get-Content -LiteralPath $Path -Encoding Byte -TotalCount 2 -ErrorAction Stop
    if ($bytes.Count -lt 2 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        throw "Downloaded payload is not a Windows executable (no MZ header): $Path"
    }
}


function Get-SqlExpressMedia {
    <#
    .SYNOPSIS
        Returns the local media path and the version stamped on the media box.
    .DESCRIPTION
        The media box carries no version in its file name and the Download
        Center exposes none over HTTP, so the version comes from the file
        version resource of the downloaded box (for example 16.0.1000.6). The
        download is cached: the Express baseline changes only when Microsoft
        republishes the package.
    #>
    param([switch]$Quiet)

    Initialize-Folder -Path $BaseDownloadRoot

    $localMedia = Join-Path $BaseDownloadRoot $MediaFileName

    if (-not (Test-Path -LiteralPath $localMedia)) {
        Write-Log "Media URL                    : $MediaUrl" -Quiet:$Quiet
        Write-Log "Downloading Express media..." -Quiet:$Quiet
        Invoke-DownloadWithRetry -Url $MediaUrl -OutFile $localMedia -Quiet:$Quiet
    }
    else {
        Write-Log "Local media exists. Skipping download." -Quiet:$Quiet
    }

    Assert-PayloadIsExecutable -Path $localMedia

    $version = (Get-Item -LiteralPath $localMedia).VersionInfo.ProductVersion
    if ($version) { $version = $version.Trim() }
    if ($version -notmatch '^\d+(\.\d+)+$') {
        throw "Media box carries no usable product version: '$version'"
    }

    Write-Log "Media product version        : $version" -Quiet:$Quiet

    return [pscustomobject]@{
        Version = $version
        Path    = $localMedia
    }
}


function New-SqlExpressWrapperContent {
    <#
    .SYNOPSIS
        Returns install and uninstall .ps1 content for the Express media box.
    .DESCRIPTION
        The box is a self-extracting archive around setup.exe. Extracting first
        and calling setup.exe directly keeps the parameter set under this
        script's control and leaves a readable setup log tree behind on the
        client. The extraction target sits at the drive root because setup
        rejects paths that contain spaces.
    #>
    param(
        [Parameter(Mandatory)][string]$MediaFile,
        [Parameter(Mandatory)][string]$Instance,
        [Parameter(Mandatory)][string]$SysAdmins
    )

    $preamble = @(
        ('$mediaPath = Join-Path $PSScriptRoot ''{0}''' -f ($MediaFile -replace "'", "''")),
        '$extractRoot = Join-Path $env:SystemDrive ''SQLEXPR_Media''',
        'if (Test-Path -LiteralPath $extractRoot) { Remove-Item -LiteralPath $extractRoot -Recurse -Force -ErrorAction SilentlyContinue }',
        '$box = Start-Process -FilePath $mediaPath -ArgumentList @(''/Q'', ("/X:" + $extractRoot)) -Wait -PassThru -NoNewWindow',
        'if ($box.ExitCode -ne 0) { exit $box.ExitCode }',
        '$setup = Join-Path $extractRoot ''setup.exe''',
        'if (-not (Test-Path -LiteralPath $setup)) { exit 1 }'
    )

    $cleanup = @(
        '$code = $proc.ExitCode',
        'Remove-Item -LiteralPath $extractRoot -Recurse -Force -ErrorAction SilentlyContinue',
        'exit $code'
    )

    $installArgs = @(
        '''/Q''',
        '''/ACTION=Install''',
        '''/IACCEPTSQLSERVERLICENSETERMS''',
        '''/SUPPRESSPRIVACYSTATEMENTNOTICE''',
        '''/FEATURES=SQLENGINE''',
        ('''/INSTANCENAME={0}''' -f $Instance),
        ('''/SQLSYSADMINACCOUNTS="{0}"''' -f $SysAdmins),
        '''/SQLSVCSTARTUPTYPE=Automatic''',
        '''/BROWSERSVCSTARTUPTYPE=Disabled''',
        '''/UPDATEENABLED=False'''
    ) -join ', '

    $uninstallArgs = @(
        '''/Q''',
        '''/ACTION=Uninstall''',
        '''/FEATURES=SQLENGINE''',
        ('''/INSTANCENAME={0}''' -f $Instance)
    ) -join ', '

    $install = ($preamble + @(
        ('$proc = Start-Process -FilePath $setup -ArgumentList @({0}) -Wait -PassThru -NoNewWindow' -f $installArgs)
    ) + $cleanup) -join "`r`n"

    $uninstall = ($preamble + @(
        ('$proc = Start-Process -FilePath $setup -ArgumentList @({0}) -Wait -PassThru -NoNewWindow' -f $uninstallArgs)
    ) + $cleanup) -join "`r`n"

    return @{
        Install   = $install
        Uninstall = $uninstall
    }
}


# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageSqlExpress {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SQL Server 2022 Express (x64) - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    $media   = Get-SqlExpressMedia
    $version = $media.Version

    Write-Log "Version                      : $version"
    Write-Log "Instance name                : $InstanceName"
    Write-Log "Sysadmin accounts            : $SysAdminAccounts"
    Write-Log ""

    # --- Versioned local content folder ---
    $localContentPath = Join-Path $BaseDownloadRoot $version
    Initialize-Folder -Path $localContentPath

    $stagedMedia = Join-Path $localContentPath $MediaFileName
    if (-not (Test-Path -LiteralPath $stagedMedia)) {
        Copy-Item -LiteralPath $media.Path -Destination $stagedMedia -Force -ErrorAction Stop
        Write-Log "Copied media to staged folder: $stagedMedia"
    }
    else {
        Write-Log "Staged media exists. Skipping copy."
    }

    # --- Generate content wrappers ---
    $wrapperContent = New-SqlExpressWrapperContent -MediaFile $MediaFileName -Instance $InstanceName -SysAdmins $SysAdminAccounts
    Write-ContentWrappers -OutputPath $localContentPath `
        -InstallPs1Content $wrapperContent.Install `
        -UninstallPs1Content $wrapperContent.Uninstall

    # --- Write stage manifest ---
    # Setup stamps the installed build on the instance Setup key; the ARP
    # entries for a SQL Server instance are per-feature and carry the media
    # baseline rather than the patched build.
    $detectionKey = "SOFTWARE\Microsoft\Microsoft SQL Server\{0}.{1}\Setup" -f $InstanceKeyPrefix, $InstanceName

    Write-Log ""
    Write-Log "Detection key                : $detectionKey"
    Write-Log "Detection value              : Version >= $version"
    Write-Log ""

    $manifestPath = Join-Path $localContentPath "stage-manifest.json"
    Write-StageManifest -Path $manifestPath -ManifestData @{
        AppName         = "Microsoft SQL Server 2022 Express"
        Publisher       = "Microsoft Corporation"
        SoftwareVersion = $version
        InstallerFile   = $MediaFileName
        InstallerType   = "EXE"
        InstallArgs     = ("/Q /ACTION=Install /IACCEPTSQLSERVERLICENSETERMS /SUPPRESSPRIVACYSTATEMENTNOTICE /FEATURES=SQLENGINE /INSTANCENAME={0} /SQLSYSADMINACCOUNTS=`"{1}`" /SQLSVCSTARTUPTYPE=Automatic /BROWSERSVCSTARTUPTYPE=Disabled /UPDATEENABLED=False" -f $InstanceName, $SysAdminAccounts)
        UninstallArgs   = ("/Q /ACTION=Uninstall /FEATURES=SQLENGINE /INSTANCENAME={0}" -f $InstanceName)
        RunningProcess  = @()
        Detection       = @{
            Type                = "RegistryKeyValue"
            RegistryKeyRelative = $detectionKey
            ValueName           = "Version"
            PropertyType        = "Version"
            Operator            = "GreaterEquals"
            ExpectedValue       = $version
            Is64Bit             = $true
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

function Invoke-PackageSqlExpress {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SQL Server 2022 Express (x64) - PACKAGE phase"
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
    Write-Log "Detection Key                : $($manifest.Detection.RegistryKeyRelative)"
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
        $media = Get-SqlExpressMedia -Quiet
        if (-not $media) { exit 1 }
        Write-Output $media.Version
        exit 0
    }
    catch {
        [Console]::Error.WriteLine("SQL Server 2022 Express GetLatestVersionOnly failed: $($_.Exception.Message)")
        exit 1
    }
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SQL Server 2022 Express (x64) Auto-Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN,$env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "MediaUrl                     : $MediaUrl"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageSqlExpress
    }
    elseif ($PackageOnly) {
        Invoke-PackageSqlExpress
    }
    else {
        Invoke-StageSqlExpress
        Invoke-PackageSqlExpress
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-sqlserver2022express'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
