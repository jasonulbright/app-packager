#Requires -Version 5.1

<#
.SYNOPSIS
    Stages or packages one application through a workbench profile, without
    the GUI.

.DESCRIPTION
    Creates a run snapshot from the selected application and profile, sets
    the child bridge environment variables, and launches the packager with
    powershell.exe exactly as the GUI does.

    Precedence of explicit parameters over the profile:

      Global defaults -> packager defaults -> profile -> variant override
      -> target override -> run override

    Every parameter listed under "Run overrides" below is a run override: it
    outranks the saved profile for this build only, is recorded in the run
    snapshot and in the build record, and is never written back into the
    profile. Omit a run override to inherit the profile's value; the profile
    in turn inherits anything it does not set. -Target and -DownloadRoot are
    run inputs rather than profile fields and always apply as given.

    Environment passed to the child:
      APP_PACKAGER_RUN_SNAPSHOT   path of this run's snapshot
      APP_PACKAGER_SIGNING        signing policy JSON
      APP_PACKAGER_WORKBENCH_ROOT data root, so the child resolves the same
      APP_PACKAGER_REQUIREMENTS / _VARIANTS / _COMMANDS / _INSTALL_MODE /
      _TITLE_MODE                 legacy bridge for consumers not yet reading
                                  the snapshot

.EXAMPLE
    .\Invoke-AppPackagerBuild.ps1 -Application catalog:package-7zip -Profile Managed -Stage

    WSUS targets:
      WSUSOnly     Package stages the application and publishes its
                   installer to WSUS; no site, share, or console.
      MECMAndWSUS  Package creates the ConfigMgr application, then
                   publishes the installer to WSUS.
    The WSUS publish reads the manifest staged in this run, so a MECMAndWSUS
    publish needs -Stage and -Package together. -Wsus* parameters outrank
    the Wsus section of the preferences file. When the approval or a decline
    fails after a successful publish, the update stays published, Wsus.Warnings
    names the failure, and the exit code is 1.

.EXAMPLE
    .\Invoke-AppPackagerBuild.ps1 -Application package-git.ps1 -Profile default -Target MECM -Package -EstimatedMinutes 10 -MaximumMinutes 25

.EXAMPLE
    .\Invoke-AppPackagerBuild.ps1 -Application package-7zip.ps1 -Target WSUSOnly -Package -WsusServer wsus01.contoso.com -WsusPort 8531 -WsusUseSsl -WsusApprovalGroup Pilot
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Application,

    [string]$Profile = 'default',

    [string]$Version,

    [ValidateSet('ContentOnly', 'MECM', 'MECMAndIntune', 'IntuneOnly', 'MECMAndWSUS', 'WSUSOnly')]
    [string]$Target = 'MECM',

    [switch]$Stage,

    [switch]$Package,

    [string]$DownloadRoot,

    [int]$EstimatedMinutes,

    [int]$MaximumMinutes,

    [string]$PackagersRoot,

    [string]$LogFolder,

    [string]$SiteCode,

    [string]$ProviderMachineName,

    [string]$FileServerPath,

    [AllowEmptyString()][string]$Comment = '',

    [string]$WsusServer,

    [ValidateRange(1, 65535)]
    [int]$WsusPort,

    [switch]$WsusUseSsl,

    [string]$WsusClassification,

    [string]$WsusApprovalGroup,

    [switch]$WsusDeclineSuperseded
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $Stage -and -not $Package) {
    throw 'Specify -Stage, -Package, or both.'
}
$publishesToWsus = ($Target -in @('MECMAndWSUS', 'WSUSOnly'))
$cliBoundParameters = @{} + $PSBoundParameters
if ($Target -eq 'MECMAndWSUS' -and $Package -and -not $Stage) {
    throw 'A MECMAndWSUS package publishes the build staged in the same run; pass -Stage together with -Package.'
}
if ([string]::IsNullOrWhiteSpace($PackagersRoot)) { $PackagersRoot = Join-Path $PSScriptRoot 'Packagers' }

# The snapshot's signing gate resolves the certificate through Get-Command, so the signing service must be
# loaded in this host before New-RunSnapshot runs or every signed build fails as "not loaded".
Import-Module (Join-Path $PackagersRoot 'AppPackagerSigning.psd1') -Force -ErrorAction Stop
Import-Module (Join-Path $PackagersRoot 'AppPackagerWorkbench.psd1') -Force -ErrorAction Stop

function Read-CliPreferences {
    $path = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'AppPackager\AppPackager.preferences.json'
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try { return ([System.IO.File]::ReadAllText($path) | ConvertFrom-Json) }
    catch { return $null }
}

function Get-CliCustomScriptRoot {
    # User-authored packagers outside the install tree live beside the
    # workbench data; the picker discovers them under the same folder.
    return (Join-Path (Get-WorkbenchDataRoot) 'scripts')
}

function Resolve-CliApplication {
    param([string]$Value, [string]$Root)

    if ($Value -match '^(catalog|custom|byo):') { return $Value }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Value)
    foreach ($candidate in @("$base.ps1", "$base.notps1")) {
        if (Test-Path -LiteralPath (Join-Path $Root $candidate)) { return (New-ApplicationId -Kind Catalog -Name $base) }
    }
    foreach ($candidate in @("$base.ps1", "$base.notps1")) {
        if (Test-Path -LiteralPath (Join-Path (Get-CliCustomScriptRoot) $candidate)) { return (New-ApplicationId -Kind Custom -Name $base) }
    }
    return (New-ApplicationId -Kind Catalog -Name $base)
}

function Resolve-CliProfile {
    param([string]$ApplicationId, [string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value) -or $Value -eq 'default') { return 'default' }
    $profiles = @(Get-Profiles -ApplicationId $ApplicationId)
    $match = @($profiles | Where-Object { $_.ProfileId -eq $Value })
    if ($match.Count -eq 0) { $match = @($profiles | Where-Object { $_.Name -eq $Value }) }
    if ($match.Count -eq 0) { throw "Profile '$Value' was not found for application '$ApplicationId'." }
    if ($match.Count -gt 1) { throw "Profile name '$Value' is ambiguous for application '$ApplicationId'; pass the profile id." }
    return [string]$match[0].ProfileId
}

function ConvertTo-CliProcessArgument {
    param([AllowNull()][string]$Argument)

    if ($null -eq $Argument -or $Argument.Length -eq 0) { return '""' }
    if ($Argument -notmatch '[\s"]') { return $Argument }
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')
    $backslashes = 0
    foreach ($character in $Argument.ToCharArray()) {
        if ($character -eq '\') { $backslashes++; continue }
        if ($character -eq '"') {
            if ($backslashes -gt 0) { [void]$builder.Append(('\' * ($backslashes * 2))) }
            [void]$builder.Append('\"')
            $backslashes = 0
            continue
        }
        if ($backslashes -gt 0) { [void]$builder.Append(('\' * $backslashes)); $backslashes = 0 }
        [void]$builder.Append($character)
    }
    if ($backslashes -gt 0) { [void]$builder.Append(('\' * ($backslashes * 2))) }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function Set-CliArgumentList {
    param(
        [Parameter(Mandatory)][System.Diagnostics.ProcessStartInfo]$StartInfo,
        [Parameter(Mandatory)][AllowEmptyString()][string[]]$Arguments
    )
    if ($StartInfo.GetType().GetProperty('ArgumentList')) {
        try { $StartInfo.ArgumentList.Clear() } catch { }
        foreach ($argument in $Arguments) { [void]$StartInfo.ArgumentList.Add($argument) }
        return
    }
    $StartInfo.Arguments = (($Arguments | ForEach-Object { ConvertTo-CliProcessArgument $_ }) -join ' ')
}

function ConvertTo-CliRequirementsJson {
    param($Profile)
    $rules = @()
    foreach ($operation in @($Profile.Requirements.Operations)) {
        if ($null -eq $operation) { continue }
        if ([string]$operation.Op -eq 'Remove') { continue }
        if (@($operation.AppliesTo) -and (@($operation.AppliesTo) -notcontains '*')) { continue }
        if ($operation.Rule) { $rules += $operation.Rule }
    }
    if ($rules.Count -eq 0) { return '' }
    return (@{ SchemaVersion = 1; Rules = $rules } | ConvertTo-Json -Depth 6 -Compress)
}

$preferences = Read-CliPreferences
$applicationId = Resolve-CliApplication -Value $Application -Root $PackagersRoot
$profileId = Resolve-CliProfile -ApplicationId $applicationId -Value $Profile
$profileObject = Get-Profile -ApplicationId $applicationId -ProfileId $profileId

# A BYO application has no packager script: it stages from the installer
# revision persisted with its definition, through the same ad-hoc path the
# drop intake uses.
$isByo = ($applicationId -like 'byo:*')
$packagerPath = $null
$byoSource = $null

if ($isByo) {
    Import-Module (Join-Path $PackagersRoot 'AppPackagerCommon.psd1') -Force -ErrorAction Stop
    $byoDefinition = Get-ApplicationDefinition -ApplicationId $applicationId
    if (-not $byoDefinition.Persisted) { throw "BYO application '$applicationId' is not stored in this data root." }
    $revisions = @(@($byoDefinition.Sources) | Where-Object { $null -ne $_ })
    if ($revisions.Count -eq 0) { throw "BYO application '$applicationId' has no stored source revision." }
    if ($PSBoundParameters.ContainsKey('Version') -and -not [string]::IsNullOrWhiteSpace($Version)) {
        $byoSource = @($revisions | Where-Object { [string]$_.SoftwareVersion -eq $Version } | Select-Object -Last 1)[0]
        if (-not $byoSource) { throw "BYO application '$applicationId' has no source revision for version '$Version'." }
    }
    else {
        $byoSource = $revisions[-1]
    }
    if (-not (Test-Path -LiteralPath ([string]$byoSource.Path) -PathType Leaf)) {
        throw "BYO source revision $($byoSource.Revision) is missing its installer: $($byoSource.Path)"
    }
}
else {
    $packagerName = ($applicationId -split ':', 2)[1]
    $scriptRoot = if ($applicationId -like 'custom:*') { Get-CliCustomScriptRoot } else { $PackagersRoot }
    foreach ($candidate in @("$packagerName.ps1", "$packagerName.notps1")) {
        $probe = Join-Path $scriptRoot $candidate
        if (Test-Path -LiteralPath $probe) { $packagerPath = $probe; break }
    }
    if (-not $packagerPath) { throw "No packager script for application '$applicationId' under $scriptRoot." }
}

if ([string]::IsNullOrWhiteSpace($DownloadRoot)) {
    $DownloadRoot = if ($preferences -and $preferences.DownloadRoot) { [string]$preferences.DownloadRoot } else { 'C:\temp\ap' }
}
$effectiveDownloadRoot = $DownloadRoot
if ($profileId -ne 'default') {
    $effectiveDownloadRoot = Join-Path (Join-Path $DownloadRoot 'profiles') $profileId
}
if ([string]::IsNullOrWhiteSpace($LogFolder)) { $LogFolder = Join-Path $DownloadRoot '_logs' }
if (-not (Test-Path -LiteralPath $LogFolder)) { New-Item -ItemType Directory -Path $LogFolder -Force | Out-Null }

$runOverrides = @{}
if ($PSBoundParameters.ContainsKey('EstimatedMinutes')) { $runOverrides['EstimatedMinutes'] = $EstimatedMinutes }
if ($PSBoundParameters.ContainsKey('MaximumMinutes')) { $runOverrides['MaximumMinutes'] = $MaximumMinutes }
if ($PSBoundParameters.ContainsKey('Version') -and -not [string]::IsNullOrWhiteSpace($Version)) { $runOverrides['PinnedVersion'] = $Version }

$signingPolicy = $null
if ($preferences -and $preferences.PSObject.Properties['ScriptSigning']) { $signingPolicy = $preferences.ScriptSigning }
$signingJson = if ($signingPolicy) { ($signingPolicy | ConvertTo-Json -Depth 6 -Compress) } else { '' }

$snapshot = New-RunSnapshot -ApplicationId $applicationId -ProfileId $profileId -Target $Target `
    -RunOverrides $runOverrides -SigningPolicy $signingPolicy -PackagerScriptPath $packagerPath -DownloadRoot $effectiveDownloadRoot

$legacy = @{
    APP_PACKAGER_RUN_SNAPSHOT   = [string]$snapshot.Path
    APP_PACKAGER_WORKBENCH_ROOT = (Get-WorkbenchDataRoot)
}
if ($signingJson) { $legacy['APP_PACKAGER_SIGNING'] = $signingJson }

$requirementsJson = ConvertTo-CliRequirementsJson -Profile $profileObject
if ($requirementsJson) { $legacy['APP_PACKAGER_REQUIREMENTS'] = $requirementsJson }
if ($profileObject.Variants -and $profileObject.Variants.Split) {
    $legacy['APP_PACKAGER_VARIANTS'] = ($profileObject.Variants.Split | ConvertTo-Json -Depth 6 -Compress)
}
# Inherit-everything profile sections are emitted as property-less objects; StrictMode makes a
# direct member access on them terminate, so every optional field is probed by name first.
function Get-CliProfileField {
    param($Container, [string]$Name)
    if ($null -eq $Container) { return $null }
    if (-not $Container.PSObject.Properties[$Name]) { return $null }
    return $Container.$Name
}

$commands = @{ SchemaVersion = 1 }
$installCommand = Get-CliProfileField (Get-CliProfileField $profileObject 'Install') 'Command'
$uninstallCommand = Get-CliProfileField (Get-CliProfileField $profileObject 'Uninstall') 'Command'
if ($installCommand) { $commands['Install'] = [string]$installCommand }
if ($uninstallCommand) { $commands['Uninstall'] = [string]$uninstallCommand }
if ($commands.Count -gt 1) { $legacy['APP_PACKAGER_COMMANDS'] = ($commands | ConvertTo-Json -Depth 3 -Compress) }
$installModeValue = Get-CliProfileField $profileObject 'InstallMode'
if ($installModeValue) { $legacy['APP_PACKAGER_INSTALL_MODE'] = [string]$installModeValue }
$titleModeValue = [string](Get-CliProfileField (Get-CliProfileField $profileObject 'Application') 'TitleMode')
$titleModeValue = switch ($titleModeValue) { 'Include version' { 'IncludeVersion' } 'No version' { 'NoVersion' } default { $titleModeValue } }
if (-not $titleModeValue -and [bool](Get-CliProfileField $preferences 'IncludeVersionInTitle')) { $titleModeValue = 'IncludeVersion' }
if (-not $titleModeValue -and [bool](Get-CliProfileField $preferences 'RemoveVersionFromTitle')) { $titleModeValue = 'NoVersion' }
if ($titleModeValue) { $legacy['APP_PACKAGER_TITLE_MODE'] = $titleModeValue }

function Invoke-CliPackagerPhase {
    param([Parameter(Mandatory)][ValidateSet('Stage', 'Package')][string]$Phase)

    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $base = [System.IO.Path]::GetFileNameWithoutExtension($packagerPath)
    $structuredLog = Join-Path $LogFolder ('{0}-{1}-{2}.structured.log' -f $base, $Phase.ToLowerInvariant(), $stamp)

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = 'powershell.exe'
    $startInfo.WorkingDirectory = Split-Path -Parent $packagerPath
    $siteFree = ($Target -in @('IntuneOnly', 'WSUSOnly'))
    $phaseSwitch = if ($Phase -eq 'Stage' -or $siteFree) { '-StageOnly' } else { '-PackageOnly' }
    $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $packagerPath, $phaseSwitch, '-LogPath', $structuredLog)

    if ($Phase -eq 'Package' -and -not $siteFree) {
        $resolvedSite = if ($SiteCode) { $SiteCode } elseif ($preferences -and $preferences.SiteCode) { [string]$preferences.SiteCode } else { 'MCM' }
        $arguments += @('-SiteCode', $resolvedSite, '-Comment', $Comment)
        $resolvedShare = if ($FileServerPath) { $FileServerPath } elseif ($preferences -and $preferences.FileShareRoot) { [string]$preferences.FileShareRoot } else { '' }
        if ($resolvedShare) {
            $head = (Get-Content -LiteralPath $packagerPath -TotalCount 120 -ErrorAction SilentlyContinue | Out-String)
            if ($head -match '\$FileServerPath') { $arguments += @('-FileServerPath', $resolvedShare) }
        }
    }
    $arguments += @('-DownloadRoot', $effectiveDownloadRoot)
    if ($runOverrides.ContainsKey('EstimatedMinutes')) {
        $arguments += @('-EstimatedRuntimeMins', [string]$runOverrides['EstimatedMinutes'])
    }
    if ($runOverrides.ContainsKey('MaximumMinutes')) {
        $arguments += @('-MaximumRuntimeMins', [string]$runOverrides['MaximumMinutes'])
    }

    Set-CliArgumentList -StartInfo $startInfo -Arguments $arguments
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    foreach ($key in $legacy.Keys) { $startInfo.EnvironmentVariables[$key] = [string]$legacy[$key] }
    if ($ProviderMachineName) { $startInfo.EnvironmentVariables['APP_PACKAGER_CM_PROVIDER'] = $ProviderMachineName }
    elseif ($preferences -and $preferences.ProviderMachineName) { $startInfo.EnvironmentVariables['APP_PACKAGER_CM_PROVIDER'] = [string]$preferences.ProviderMachineName }
    if ($preferences -and [string]$preferences.ContentFallback) { $startInfo.EnvironmentVariables['APP_PACKAGER_DT_CONTENT_FALLBACK'] = [string]$preferences.ContentFallback }
    if ($preferences -and [string]$preferences.SlowNetworkDeploymentMode) { $startInfo.EnvironmentVariables['APP_PACKAGER_DT_SLOW_NETWORK_MODE'] = [string]$preferences.SlowNetworkDeploymentMode }

    $process = [System.Diagnostics.Process]::Start($startInfo)
    $process.WaitForExit()
    return [pscustomobject]@{ Phase = $Phase; ExitCode = $process.ExitCode; StructuredLog = $structuredLog }
}

function Invoke-CliAdHocPhases {
    <#
        The ad-hoc stage and package run in this process, so the child bridge
        variables sit on the process environment for the duration of the run
        only: leaving them behind would hand the next build a policy and a
        data root nobody selected for it.
    #>
    $saved = @{}
    foreach ($key in $legacy.Keys) {
        $saved[$key] = [Environment]::GetEnvironmentVariable($key, 'Process')
        [Environment]::SetEnvironmentVariable($key, [string]$legacy[$key], 'Process')
    }
    $phases = New-Object System.Collections.ArrayList
    try {
        $analysis = Get-InstallerAnalysis -Path ([string]$byoSource.Path)
        $stageResult = $null
        if ($Package -and -not $Stage -and -not $publishesToWsus) {
            # Package alone consumes the build the last -Stage left behind; the
            # folder shape mirrors the ad-hoc stage so nothing is re-staged. A
            # WSUS target stages again, because its publish accepts only a
            # manifest that carries this run's BuildId.
            $sanitize = { param($s) (($s -replace '[\\/:*?"<>|]', '') -replace '\s+', ' ').Trim() }
            $vendorFolder = & $sanitize ([string]$byoDefinition.Publisher)
            if ([string]::IsNullOrWhiteSpace($vendorFolder)) { $vendorFolder = 'Unknown' }
            $appFolder = & $sanitize ([string]$byoDefinition.DisplayName)
            $versionFolder = & $sanitize ([string]$byoSource.SoftwareVersion)
            $stagedPath = Join-Path (Join-Path (Join-Path $effectiveDownloadRoot $vendorFolder) $appFolder) $versionFolder
            $manifestPath = Join-Path $stagedPath 'stage-manifest.json'
            if (Test-Path -LiteralPath $manifestPath) {
                $stageResult = [pscustomobject]@{ StagedPath = $stagedPath; ManifestPath = $manifestPath; VendorFolder = $vendorFolder; AppFolder = $appFolder }
            }
        }
        if ($Stage -or ($Package -and -not $stageResult)) {
            $stageArguments = @{
                Analysis        = $analysis
                DownloadRoot    = $effectiveDownloadRoot
                AppName         = [string]$byoDefinition.DisplayName
                Publisher       = [string]$byoDefinition.Publisher
                SoftwareVersion = [string]$byoSource.SoftwareVersion
            }
            # The ad-hoc stage emits progress objects ahead of its result, so the pipeline is filtered
            # down to the record that actually carries the staged path.
            $stageResult = @(New-AdHocStage @stageArguments | Where-Object { $_ -and $_.PSObject.Properties['StagedPath'] })[-1]
            [void]$phases.Add([pscustomobject]@{ Phase = 'Stage'; ExitCode = 0; StagedPath = $stageResult.StagedPath })
        }
        if ($Package -and $Target -notin @('ContentOnly', 'IntuneOnly', 'WSUSOnly')) {
            $resolvedSite = if ($SiteCode) { $SiteCode } elseif ($preferences -and $preferences.SiteCode) { [string]$preferences.SiteCode } else { 'MCM' }
            $resolvedProvider = if ($ProviderMachineName) { $ProviderMachineName } elseif ($preferences -and $preferences.ProviderMachineName) { [string]$preferences.ProviderMachineName } else { '' }
            $resolvedShare = if ($FileServerPath) { $FileServerPath } elseif ($preferences -and $preferences.FileShareRoot) { [string]$preferences.FileShareRoot } else { '' }
            $layout = if ($preferences -and $preferences.ContentLayout) { [string]$preferences.ContentLayout } else { 'Nested' }
            if (-not (Connect-CMSite -SiteCode $resolvedSite -ProviderMachineName $resolvedProvider)) {
                throw "Site connection to '$resolvedSite' failed; the staged content is intact and packaging was skipped."
            }
            # Connecting leaves the session on the CM provider drive, where the ad-hoc packager's bare UNC
            # content paths resolve through the provider and collapse to a null path.
            Set-Location -LiteralPath $PSScriptRoot
            $packageArguments = @{
                StagedPath      = $stageResult.StagedPath
                VendorFolder    = $stageResult.VendorFolder
                AppFolder       = $stageResult.AppFolder
                FileServerPath  = $resolvedShare
                SiteCode        = $resolvedSite
                Comment         = $Comment
                ContentLayout   = $layout
            }
            if ($runOverrides.ContainsKey('EstimatedMinutes')) { $packageArguments['EstimatedRuntimeMins'] = [int]$runOverrides['EstimatedMinutes'] }
            if ($runOverrides.ContainsKey('MaximumMinutes')) { $packageArguments['MaximumRuntimeMins'] = [int]$runOverrides['MaximumMinutes'] }
            [void](Invoke-AdHocPackage @packageArguments)
            [void]$phases.Add([pscustomobject]@{ Phase = 'Package'; ExitCode = 0; StagedPath = $stageResult.StagedPath })
        }
    }
    finally {
        foreach ($key in $saved.Keys) { [Environment]::SetEnvironmentVariable($key, $saved[$key], 'Process') }
    }
    return @($phases)
}

$results = New-Object System.Collections.ArrayList
if ($isByo) {
    foreach ($phase in (Invoke-CliAdHocPhases)) { [void]$results.Add($phase) }
}
else {
    if ($Stage) { [void]$results.Add((Invoke-CliPackagerPhase -Phase 'Stage')) }
    if ($Package -and @(@($results) | Where-Object { $_.ExitCode -ne 0 }).Count -eq 0) {
        [void]$results.Add((Invoke-CliPackagerPhase -Phase 'Package'))
    }
}

$exitCode = 0
foreach ($result in $results) { if ($result.ExitCode -ne 0) { $exitCode = $result.ExitCode } }

function Resolve-CliWsusSettings {
    # Explicit parameters outrank the preferences file; the publisher
    # validates every value again before it connects. A stored port belongs
    # to the stored SSL choice, so an SSL change without a port takes the
    # default port of the new scheme.
    param([Parameter(Mandatory)][hashtable]$Bound)

    $stored = @{}
    if ($preferences -and $preferences.PSObject.Properties['Wsus'] -and $preferences.Wsus) {
        foreach ($property in $preferences.Wsus.PSObject.Properties) { $stored[$property.Name] = $property.Value }
    }
    $settings = @{
        ServerName        = [string]$stored['ServerName']
        UseSsl            = [bool]$stored['UseSsl']
        Classification    = $(if ($stored['Classification']) { [string]$stored['Classification'] } else { 'Updates' })
        ApprovalGroup     = [string]$stored['ApprovalGroup']
        DeclineSuperseded = [bool]$stored['DeclineSuperseded']
    }
    if ($Bound.ContainsKey('WsusServer'))            { $settings.ServerName = [string]$Bound['WsusServer'] }
    if ($Bound.ContainsKey('WsusUseSsl'))            { $settings.UseSsl = [bool]$Bound['WsusUseSsl'] }
    if ($Bound.ContainsKey('WsusClassification'))    { $settings.Classification = [string]$Bound['WsusClassification'] }
    if ($Bound.ContainsKey('WsusApprovalGroup'))     { $settings.ApprovalGroup = [string]$Bound['WsusApprovalGroup'] }
    if ($Bound.ContainsKey('WsusDeclineSuperseded')) { $settings.DeclineSuperseded = [bool]$Bound['WsusDeclineSuperseded'] }

    $storedPort = 0
    [void][int]::TryParse([string]$stored['PortNumber'], [ref]$storedPort)
    if ($Bound.ContainsKey('WsusPort')) { $settings['PortNumber'] = [int]$Bound['WsusPort'] }
    elseif ($storedPort -ge 1 -and [bool]$stored['UseSsl'] -eq $settings.UseSsl) { $settings['PortNumber'] = $storedPort }
    else { $settings['PortNumber'] = $(if ($settings.UseSsl) { 8531 } else { 8530 }) }
    return $settings
}

$wsusOutcome = $null
if ($publishesToWsus -and $Package -and $exitCode -eq 0) {
    try {
        Import-Module (Join-Path $PackagersRoot 'AppPackagerCommon.psd1') -ErrorAction Stop
        Import-Module (Join-Path $PackagersRoot 'AppPackagerWsus.psd1') -Force -ErrorAction Stop
        $wsusSettings = Resolve-CliWsusSettings -Bound $cliBoundParameters
        if ([string]::IsNullOrWhiteSpace([string]$wsusSettings.ServerName)) {
            throw 'No WSUS server: pass -WsusServer or set the server in Options, WSUS Publishing.'
        }
        $staged = Resolve-StageManifestForBuild -BuildId ([string]$snapshot.BuildId) -SearchRoot $effectiveDownloadRoot
        $stagedManifest = Read-StageManifest -Path ([string]$staged.Path)
        $published = Publish-WsusSoftwareUpdate -Manifest $stagedManifest -ContentFolder ([string]$staged.StageRoot) -Settings $wsusSettings
        $wsusOutcome = [pscustomobject]@{ Ok = $true; PackageId = [string]$published.PackageId; Outcome = [string]$published.Outcome; Warnings = @($published.Warnings); Message = [string]$published.Message }
        if (@($published.Warnings).Count -gt 0) { $exitCode = 1 }
    }
    catch {
        $wsusOutcome = [pscustomobject]@{ Ok = $false; PackageId = ''; Outcome = 'Failed'; Warnings = @(); Message = $_.Exception.Message }
        $exitCode = 1
    }
}

[pscustomobject]@{
    ApplicationId = $applicationId
    ProfileId     = $profileId
    BuildId       = [string]$snapshot.BuildId
    Snapshot      = [string]$snapshot.Path
    Target        = $Target
    DownloadRoot  = $effectiveDownloadRoot
    Phases        = @($results)
    Wsus          = $wsusOutcome
    ExitCode      = $exitCode
}

exit $exitCode
