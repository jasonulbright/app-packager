#Requires -Version 5.1
<#
.SYNOPSIS
    Converts an existing ConfigMgr application to Intune Win32 apps.

.DESCRIPTION
    Reads the application from the site (read-only), plans one Intune Win32 app
    per deployment type, copies the content, writes a conversion report, and
    can build the .intunewin files and publish them through Microsoft Graph v1.0.

    ConfigMgr installs the first deployment type whose requirements pass.
    Intune has no deployment types, so each deployment type becomes its own
    app. With the default PerDeploymentType strategy, the requirement rules of
    each app keep the "first match wins" meaning. With AssignByGroup, the apps
    carry only their own requirements and each app goes to its own group.

    Dependencies and supersedence are listed in relationships.json and are not
    created: Microsoft Graph v1.0 has no resource for them.

    Run it on a computer with the ConfigMgr console installed.

.PARAMETER ApplicationName
    The exact name of the ConfigMgr application.

.PARAMETER SiteCode
    The ConfigMgr site code.

.PARAMETER OutputRoot
    The folder that receives one folder per application.

.PARAMETER VariantStrategy
    PerDeploymentType (default) or AssignByGroup.

.PARAMETER NativeDetectionOnly
    Refuse a detection that needs a PowerShell script in Intune (Or, groups, and
    comparisons without a native rule).

.PARAMETER SkipContent
    Plan and write the reports without copying content.

.PARAMETER Package
    Also build the .intunewin files.

.PARAMETER Publish
    Also publish each app through Microsoft Graph. Needs the tenant values, or
    -UseSavedSettings.

.PARAMETER UseSavedSettings
    Read the tenant ID, the client ID and the client secret from the Intune
    settings that App Packager saved in AppPackager.preferences.json.

.EXAMPLE
    .\Convert-CMApplicationToIntune.ps1 -ApplicationName 'Contoso Viewer' -SiteCode MCM -OutputRoot D:\Converted

.EXAMPLE
    .\Convert-CMApplicationToIntune.ps1 -ApplicationName 'Contoso Viewer' -SiteCode MCM -OutputRoot D:\Converted -Publish -UseSavedSettings
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ApplicationName,
    [Parameter(Mandatory)][string]$SiteCode,
    [Parameter(Mandatory)][string]$OutputRoot,
    [ValidateSet('PerDeploymentType', 'AssignByGroup')][string]$VariantStrategy = 'PerDeploymentType',
    [switch]$NativeDetectionOnly,
    [switch]$SkipContent,
    [switch]$Package,
    [switch]$Publish,
    [switch]$UseSavedSettings,
    [string]$TenantId = '',
    [string]$ClientId = '',
    [System.Security.SecureString]$ClientSecret,
    [string]$ToolPath = ''
)

$ErrorActionPreference = 'Stop'
$packagers = Join-Path $PSScriptRoot 'Packagers'
Import-Module (Join-Path $packagers 'AppPackagerCommon.psd1') -Force -DisableNameChecking
Import-Module (Join-Path $packagers 'AppPackagerCmImport.psd1') -Force

if ($Publish -and $UseSavedSettings) {
    $prefsPath = Join-Path $PSScriptRoot 'AppPackager.preferences.json'
    if (-not (Test-Path -LiteralPath $prefsPath)) { throw "Saved settings were not found: $prefsPath" }
    $intune = (Get-Content -LiteralPath $prefsPath -Raw -Encoding UTF8 | ConvertFrom-Json).Intune
    if (-not $intune -or [string]::IsNullOrWhiteSpace([string]$intune.TenantId) -or [string]::IsNullOrWhiteSpace([string]$intune.ClientSecretProtected)) {
        throw 'The saved settings hold no Intune tenant, client or secret. Set them in Options, or pass -TenantId, -ClientId and -ClientSecret.'
    }
    $TenantId = [string]$intune.TenantId
    $ClientId = [string]$intune.ClientId
    $ClientSecret = ConvertTo-SecureString -String ([string]$intune.ClientSecretProtected)
}

$plain = ''
if ($ClientSecret) {
    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
    try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer) }
}

$params = @{
    ApplicationName = $ApplicationName
    OutputRoot      = $OutputRoot
    SiteCode        = $SiteCode
    VariantStrategy = $VariantStrategy
    ToolPath        = $ToolPath
    TenantId        = $TenantId
    ClientId        = $ClientId
    ClientSecret    = $plain
}
if ($NativeDetectionOnly) { $params['NativeDetectionOnly'] = $true }
if ($SkipContent) { $params['SkipContent'] = $true }
if ($Package) { $params['Package'] = $true }
if ($Publish) { $params['Publish'] = $true }

$result = Convert-CMApplicationToIntune @params
Format-CMImportReport -Plan $result.Plan | Write-Host
Write-Host ''
Write-Host ('Output folder: {0}' -f $result.Stage.ApplicationFolder)
foreach ($built in $result.Packages) {
    Write-Host ("{0}: {1}{2}" -f $built.Name, $built.IntuneWinPath, $(if ($built.AppId) { " (Intune app " + $built.AppId + ")" } else { "" }))
}
if (@($result.Plan.DeploymentTypes | Where-Object { $_.Convert }).Count -eq 0) {
    Write-Error 'No deployment type could be converted. The report lists the reasons.'
    exit 1
}
