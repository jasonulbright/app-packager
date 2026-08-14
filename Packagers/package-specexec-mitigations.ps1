<#
Vendor: Microsoft
App: Speculative Execution Mitigations
CMName: Speculative Execution Mitigations (Intel-AMD-BHI)
VendorUrl: https://support.microsoft.com/en-us/topic/kb4073119-windows-client-guidance-for-it-pros-to-protect-against-silicon-based-microarchitectural-and-speculative-execution-side-channel-vulnerabilities-35820a8a-ae13-1299-88cc-357f104f5b11
CPE:
ReleaseNotesUrl:
DownloadPageUrl:
UpdateCadenceDays: 0

.SYNOPSIS
    Packages the speculative-execution CVE registry mitigations as one MECM
    Application with 6 deployment types gated by global conditions.

.DESCRIPTION
    Creates a single Script-installer Application with one deployment type per
    distinct registry payload -- 6 total: {Intel HT-on, Intel HT-off, AMD} x
    {standard, Hyper-V host}. Each deployment type carries requirement rules built
    from global conditions, so a device deployed the app installs exactly the
    deployment type that matches its hardware. The full-remediation values ("both
    boxes" per vendor section in Registry\New-SpecExecBitmask.ps1) are:

      Intel, HT enabled   FeatureSettingsOverride = 0x00800048  (prior CVE bundle + BHI)
      Intel, HT disabled  FeatureSettingsOverride = 0x00802048  (prior CVE bundle + BHI)
      AMD (either HT)     FeatureSettingsOverride = 0x05000040  (BTC + Inception)
      All                 FeatureSettingsOverrideMask = 0x00000003
      Hyper-V hosts       MinVmVersionForCpuBasedMitigations = "1.0" (additional)

    Only payload differences produce deployment types: workstation vs. server never
    changes the registry values (the Hyper-V role does), so there is no OS-class
    split; the AMD override does not vary with HT state, so AMD gets no HT split.
    The 12 SpecExec device collections
    (general-scripts\MECM\Collections\New-SpecExecTargetCollections.ps1) remain the
    deployment-targeting and reporting surface; multiple collections simply resolve
    to the same deployment type. Requirement rules fully partition the fleet, so
    deployment type priority order never decides the outcome.

    Global conditions (looked up by name first; created only when absent, using
    built-in setting types wherever the data is reachable without a script):

      CPU Manufacturer          WQL query: Win32_Processor.Manufacturer (String)
      Hyper-Threading Enabled   Script (Boolean): any Win32_Processor with
                                NumberOfLogicalProcessors > NumberOfCores.
                                Script type is required here: WQL WHERE clauses
                                cannot compare two properties of the same instance.
                                Only the Intel deployment types use this rule.
      Hyper-V Role Installed    Registry key existence: SYSTEM\CurrentControlSet\
                                Services\vmms (present only with the Hyper-V role)

    Detection is registry-based per deployment type: FeatureSettingsOverride equals
    the deployment type's DWORD AND FeatureSettingsOverrideMask equals 3 (AND
    MinVmVersionForCpuBasedMitigations equals "1.0" on the Hyper-V types).

    Install writes the values and exits 3010 (mitigations activate after reboot);
    deployment types use RebootBehavior BasedOnExitCode so the pending reboot is
    honored per deployment settings / maintenance windows.

    Supports two-phase operation:
      -StageOnly    Write Install/Uninstall script content to the local staging folder
      -PackageOnly  Copy content to network, ensure global conditions, create the
                    application and its 6 deployment types

.PARAMETER SiteCode
    ConfigMgr site code PSDrive name (e.g., "MCM").

.PARAMETER Comment
    Free-form change/WO text stored on the CM Application Description field.

.PARAMETER FileServerPath
    UNC root that contains your Applications folder (example: \\fileserver\sccm$).
    Content is staged under: <FileServerPath>\Applications\Microsoft\SpecExec-Mitigations\<Version>

.PARAMETER DownloadRoot
    Local root folder for staging content. Default: C:\temp\ap

.PARAMETER AppName
    CM Application name. Default: "Speculative Execution Mitigations (Intel-AMD-BHI)".

.PARAMETER ContentVersion
    Version folder / CM SoftwareVersion. Bump when mitigation values change.
    Default: 2026.08

.PARAMETER CpuVendorGlobalConditionName
    Name of the global condition returning Win32_Processor.Manufacturer.
    Reused if it already exists. Default: "CPU Manufacturer"

.PARAMETER HtGlobalConditionName
    Name of the Boolean global condition for SMT/Hyper-Threading state.
    Reused if it already exists. Default: "Hyper-Threading Enabled"

.PARAMETER HyperVGlobalConditionName
    Name of the registry-key-existence global condition for the Hyper-V role.
    Reused if it already exists. Default: "Hyper-V Role Installed"

.PARAMETER EstimatedRuntimeMins
    Estimated runtime in minutes for each deployment type. Default: 5

.PARAMETER MaximumRuntimeMins
    Maximum allowed runtime in minutes for each deployment type. Default: 15

.EXAMPLE
    .\package-specexec-mitigations.ps1 -SiteCode MCM -FileServerPath \\fileserver\sccm$ -Comment "WO-1234"

.REQUIREMENTS
    - PowerShell 5.1
    - ConfigMgr Admin Console installed (ConfigurationManager PowerShell module available)
    - RBAC permissions to create Applications, Deployment Types, and Global Conditions
    - Write access to FileServerPath
#>

param(
    [string]$SiteCode = "MCM",
    [string]$Comment = "",
    [string]$FileServerPath = "\\fileserver\sccm$",
    [string]$DownloadRoot = "C:\temp\ap",
    [string]$AppName = "Speculative Execution Mitigations (Intel-AMD-BHI)",
    [string]$ContentVersion = "2026.08",
    [string]$CpuVendorGlobalConditionName = "CPU Manufacturer",
    [string]$HtGlobalConditionName = "Hyper-Threading Enabled",
    [string]$HyperVGlobalConditionName = "Hyper-V Role Installed",
    [int]$EstimatedRuntimeMins = 5,
    [int]$MaximumRuntimeMins = 15,
    [string]$LogPath,
    [switch]$GetLatestVersionOnly,
    [switch]$StageOnly,
    [switch]$PackageOnly,
    [switch]$VerboseLog
)


Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
Initialize-Logging -LogPath $LogPath -VerboseLogging:$VerboseLog

if ($StageOnly -and $PackageOnly) {
    Write-Log "-StageOnly and -PackageOnly cannot be used together." -Level ERROR
    exit 1
}

# --- Configuration ---
$VendorFolder = "Microsoft"
$AppFolder    = "SpecExec-Mitigations"

$BaseDownloadRoot = Join-Path $DownloadRoot "SpecExec-Mitigations"
$InstallScriptName   = "Install-SpecExec.ps1"
$UninstallScriptName = "Uninstall-SpecExec.ps1"

$MemMgmtKey = 'SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'
$HvKey      = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization'
$HvValue    = 'MinVmVersionForCpuBasedMitigations'
$MaskDec    = 3

# Full-remediation FeatureSettingsOverride values from New-SpecExecBitmask.ps1:
# Intel = prior CVE bundle (HT-dependent) OR BHI; AMD = BTC OR Inception.
$OverrideValues = @{
    IntelHtOn  = 0x00800048
    IntelHtOff = 0x00802048
    Amd        = 0x05000040
}

# ---------------------------------------------------------------------------
# Deployment type matrix
# ---------------------------------------------------------------------------

function Get-SpecExecDeploymentTypeMatrix {
    # One deployment type per distinct registry payload. HtEnabled is $null for
    # AMD (the AMD override does not vary with HT, so no HT requirement is added).
    $payloads = @(
        @{ Label = 'Intel - HT Enabled';  Vendor = 'Intel'; HtEnabled = $true;  Override = $OverrideValues.IntelHtOn }
        @{ Label = 'Intel - HT Disabled'; Vendor = 'Intel'; HtEnabled = $false; Override = $OverrideValues.IntelHtOff }
        @{ Label = 'AMD';                 Vendor = 'AMD';   HtEnabled = $null;  Override = $OverrideValues.Amd }
    )

    $matrix = foreach ($hyperV in $false, $true) {
        foreach ($p in $payloads) {
            [pscustomobject]@{
                Name        = if ($hyperV) { '{0} - Hyper-V Host' -f $p.Label } else { $p.Label }
                Vendor      = $p.Vendor
                HtEnabled   = $p.HtEnabled
                OverrideDec = [int]$p.Override
                OverrideHex = '0x{0:X8}' -f $p.Override
                IsHyperV    = $hyperV
            }
        }
    }
    return $matrix
}

# ---------------------------------------------------------------------------
# Stage phase
# ---------------------------------------------------------------------------

function Invoke-StageSpecExec {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SpecExec Mitigations - STAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    $localContentPath = Join-Path $BaseDownloadRoot $ContentVersion
    Initialize-Folder -Path $localContentPath

    $installPs1 = @(
        'param('
        '    [Parameter(Mandatory)][string]$Override,'
        '    [switch]$HyperV'
        ')'
        '$val = [Convert]::ToInt32($Override, 16)'
        "`$mm = 'HKLM:\$MemMgmtKey'"
        "New-ItemProperty -Path `$mm -Name 'FeatureSettingsOverride'     -Value `$val -PropertyType DWord -Force | Out-Null"
        "New-ItemProperty -Path `$mm -Name 'FeatureSettingsOverrideMask' -Value $MaskDec    -PropertyType DWord -Force | Out-Null"
        'if ($HyperV) {'
        "    `$hv = 'HKLM:\$HvKey'"
        '    if (-not (Test-Path -LiteralPath $hv)) { New-Item -Path $hv -Force | Out-Null }'
        "    New-ItemProperty -Path `$hv -Name '$HvValue' -Value '1.0' -PropertyType String -Force | Out-Null"
        '}'
        '# 3010 = success, reboot required (mitigations activate at next boot)'
        'exit 3010'
    ) -join "`r`n"

    $uninstallPs1 = @(
        "`$mm = 'HKLM:\$MemMgmtKey'"
        "Remove-ItemProperty -Path `$mm -Name 'FeatureSettingsOverride'     -ErrorAction SilentlyContinue"
        "Remove-ItemProperty -Path `$mm -Name 'FeatureSettingsOverrideMask' -ErrorAction SilentlyContinue"
        "Remove-ItemProperty -Path 'HKLM:\$HvKey' -Name '$HvValue' -ErrorAction SilentlyContinue"
        'exit 3010'
    ) -join "`r`n"

    Set-Content -LiteralPath (Join-Path $localContentPath $InstallScriptName)   -Value $installPs1   -Encoding ASCII -Force -ErrorAction Stop
    Write-Log "Wrote content                : $InstallScriptName"
    Set-Content -LiteralPath (Join-Path $localContentPath $UninstallScriptName) -Value $uninstallPs1 -Encoding ASCII -Force -ErrorAction Stop
    Write-Log "Wrote content                : $UninstallScriptName"

    Write-Log ""
    Write-Log "Stage complete               : $localContentPath"

    return $localContentPath
}

# ---------------------------------------------------------------------------
# Package phase helpers
# ---------------------------------------------------------------------------

function Get-OrCreateGlobalConditions {
    <#
    Returns a hashtable of the three global condition objects, reusing any that
    already exist by name. Must be called from the CM site drive.
    #>
    $conditions = @{}

    $existing = Get-CMGlobalCondition -Name $CpuVendorGlobalConditionName
    if ($existing) {
        Write-Log "Global condition (existing)  : $CpuVendorGlobalConditionName"
        $conditions.CpuVendor = $existing
    }
    else {
        Write-Log "Global condition (creating)  : $CpuVendorGlobalConditionName"
        $conditions.CpuVendor = New-CMGlobalConditionWqlQuery `
            -Name $CpuVendorGlobalConditionName `
            -DataType String `
            -Namespace 'root\cimv2' `
            -Class 'Win32_Processor' `
            -Property 'Manufacturer' `
            -Description 'Win32_Processor.Manufacturer (GenuineIntel / AuthenticAMD).'
    }

    $existing = Get-CMGlobalCondition -Name $HtGlobalConditionName
    if ($existing) {
        Write-Log "Global condition (existing)  : $HtGlobalConditionName"
        $conditions.Ht = $existing
    }
    else {
        Write-Log "Global condition (creating)  : $HtGlobalConditionName"
        # Script type: WQL WHERE clauses cannot compare NumberOfLogicalProcessors
        # against NumberOfCores (no property-to-property comparison in WMI WQL).
        $htScript = @(
            '$smt = $false'
            'Get-CimInstance -ClassName Win32_Processor | ForEach-Object {'
            '    if ($_.NumberOfLogicalProcessors -gt $_.NumberOfCores) { $smt = $true }'
            '}'
            '$smt'
        ) -join "`r`n"
        $conditions.Ht = New-CMGlobalConditionScript `
            -Name $HtGlobalConditionName `
            -DataType Boolean `
            -ScriptLanguage PowerShell `
            -ScriptText $htScript `
            -Description 'True when any Win32_Processor reports NumberOfLogicalProcessors > NumberOfCores (SMT/HT enabled).'
    }

    $existing = Get-CMGlobalCondition -Name $HyperVGlobalConditionName
    if ($existing) {
        Write-Log "Global condition (existing)  : $HyperVGlobalConditionName"
        $conditions.HyperV = $existing
    }
    else {
        Write-Log "Global condition (creating)  : $HyperVGlobalConditionName"
        $conditions.HyperV = New-CMGlobalConditionRegistryKey `
            -Name $HyperVGlobalConditionName `
            -RegistryHive LocalMachine `
            -KeyName 'SYSTEM\CurrentControlSet\Services\vmms' `
            -Is64Bit $true `
            -Description 'vmms service key exists only when the Hyper-V role is installed.'
    }

    return $conditions
}

function New-SpecExecRequirementRules {
    <#
    Builds the requirement rule array for one deployment type entry.
    Must be called from the CM site drive. Rule objects are not reused
    across deployment types; each carries its own rule id.
    #>
    param(
        [Parameter(Mandatory)][pscustomobject]$Entry,
        [Parameter(Mandatory)][hashtable]$Conditions
    )

    $vendorString = if ($Entry.Vendor -eq 'Intel') { 'GenuineIntel' } else { 'AuthenticAMD' }

    $rules = @()
    $rules += New-CMRequirementRuleCommonValue -InputObject $Conditions.CpuVendor -RuleOperator IsEquals -Value1 $vendorString
    if ($null -ne $Entry.HtEnabled) {
        $rules += New-CMRequirementRuleBooleanValue -InputObject $Conditions.Ht -Value $Entry.HtEnabled
    }
    $rules += New-CMRequirementRuleExistential -InputObject $Conditions.HyperV -Existential $Entry.IsHyperV

    return $rules
}

function New-SpecExecDetectionClauses {
    <#
    Builds the registry detection clause array for one deployment type entry
    (all clauses joined by AND, the Add-CMScriptDeploymentType default).
    Must be called from a filesystem drive, not the CM PSDrive.
    #>
    param([Parameter(Mandatory)][pscustomobject]$Entry)

    # -Is64Bit is deliberately absent: on detection clauses it means "32-bit app
    # view (WOW6432Node)", the opposite of the global-condition cmdlets' Is64Bit.
    # Both keys live in non-redirected native paths.
    $clauses = @()
    $clauses += New-CMDetectionClauseRegistryKeyValue `
        -Hive LocalMachine -KeyName $MemMgmtKey -ValueName 'FeatureSettingsOverride' `
        -PropertyType Integer -Value -ExpressionOperator IsEquals -ExpectedValue $Entry.OverrideDec
    $clauses += New-CMDetectionClauseRegistryKeyValue `
        -Hive LocalMachine -KeyName $MemMgmtKey -ValueName 'FeatureSettingsOverrideMask' `
        -PropertyType Integer -Value -ExpressionOperator IsEquals -ExpectedValue $MaskDec
    if ($Entry.IsHyperV) {
        $clauses += New-CMDetectionClauseRegistryKeyValue `
            -Hive LocalMachine -KeyName $HvKey -ValueName $HvValue `
            -PropertyType String -Value -ExpressionOperator IsEquals -ExpectedValue '1.0'
    }
    return $clauses
}

# ---------------------------------------------------------------------------
# Package phase
# ---------------------------------------------------------------------------

function Invoke-PackageSpecExec {
    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SpecExec Mitigations - PACKAGE phase"
    Write-Log ("=" * 60)
    Write-Log ""

    $localContentPath = Join-Path $BaseDownloadRoot $ContentVersion
    if (-not (Test-Path -LiteralPath (Join-Path $localContentPath $InstallScriptName))) {
        throw "Staged content not found - run Stage phase first: $localContentPath"
    }

    # --- Network share ---
    if (-not (Test-NetworkShareAccess -Path $FileServerPath)) {
        throw "Network root path not accessible: $FileServerPath"
    }

    $networkAppRoot = Get-NetworkAppRoot -FileServerPath $FileServerPath -VendorFolder $VendorFolder -AppFolder $AppFolder
    $networkContentPath = Join-Path $networkAppRoot $ContentVersion
    Initialize-Folder -Path $networkContentPath

    Write-Log "Network content path         : $networkContentPath"

    foreach ($f in (Get-ChildItem -Path $localContentPath -File -ErrorAction Stop)) {
        Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $networkContentPath $f.Name) -Force -ErrorAction Stop
        Write-Log "Copied to network            : $($f.Name)"
    }
    Write-Log ""

    $matrix = Get-SpecExecDeploymentTypeMatrix

    $step = 'initialization'
    try {
        $step = "Connect-CMSite (SiteCode=$SiteCode)"
        if (-not (Connect-CMSite -SiteCode $SiteCode)) {
            throw "CM site connection failed."
        }

        # --- Global conditions (reuse-or-create, on the CM drive) ---
        $step = 'Get-OrCreateGlobalConditions'
        $conditions = Get-OrCreateGlobalConditions

        # --- Application ---
        $step = "Get-CMApplication duplicate check ('$AppName')"
        $existing = Get-CMApplication -Name $AppName -ErrorAction SilentlyContinue
        if ($existing) {
            $existingApps = @($existing)
            if ($existingApps.Count -gt 1) {
                throw "Multiple existing MECM applications matched '$AppName'; refusing to package until the duplicate names are resolved."
            }
            $cmApp = $existingApps[0]
            $resumeDts = @(Get-CMDeploymentType -ApplicationName $AppName -ErrorAction SilentlyContinue)
            Write-Log "Application already exists   : $AppName (has $($resumeDts.Count) of $($matrix.Count) deployment types; resuming, existing ones are skipped)" -Level WARN
        }
        else {
            Write-Log "Creating CM Application      : $AppName"
            $step = "New-CMApplication ('$AppName')"
            $cmApp = New-CMApplication `
                -Name $AppName `
                -Publisher 'Microsoft' `
                -SoftwareVersion $ContentVersion `
                -Description $Comment `
                -AutoInstall $true `
                -ErrorAction Stop
            Write-Log "Application CI_ID            : $($cmApp.CI_ID)"
        }

        # --- Deployment types ---
        foreach ($entry in $matrix) {
            Write-Log ""
            Write-Log "Deployment type              : $($entry.Name)"

            # Get-CMDeploymentType directly: Test-MECMApplicationHasDeploymentType
            # is internal to AppPackagerCommon (not in FunctionsToExport).
            if (Get-CMDeploymentType -ApplicationName $AppName -DeploymentTypeName $entry.Name -ErrorAction SilentlyContinue) {
                Write-Log "  Already exists, skipping." -Level WARN
                continue
            }

            # Requirement rules are built on the CM drive; detection clauses
            # must be built off it (CM PSDrive context breaks clause binding).
            $step = "New-SpecExecRequirementRules ('$($entry.Name)')"
            $requirements = New-SpecExecRequirementRules -Entry $entry -Conditions $conditions

            $step = "New-SpecExecDetectionClauses ('$($entry.Name)')"
            Set-Location C: -ErrorAction Stop
            $clauses = New-SpecExecDetectionClauses -Entry $entry

            $step = "Connect-CMSite reconnect (SiteCode=$SiteCode)"
            if (-not (Connect-CMSite -SiteCode $SiteCode)) {
                throw "CM site reconnection failed."
            }

            $installCommand = 'PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "{0}" -Override {1}' -f $InstallScriptName, $entry.OverrideHex
            if ($entry.IsHyperV) { $installCommand += ' -HyperV' }
            $uninstallCommand = 'PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "{0}"' -f $UninstallScriptName

            $step = "Add-CMScriptDeploymentType ('$($entry.Name)')"
            Add-CMScriptDeploymentType `
                -ApplicationName $AppName `
                -DeploymentTypeName $entry.Name `
                -ContentLocation $networkContentPath `
                -InstallCommand $installCommand `
                -UninstallCommand $uninstallCommand `
                -AddDetectionClause $clauses `
                -AddRequirement $requirements `
                -InstallationBehaviorType InstallForSystem `
                -LogonRequirementType WhetherOrNotUserLoggedOn `
                -UserInteractionMode Hidden `
                -RebootBehavior BasedOnExitCode `
                -EstimatedRuntimeMins $EstimatedRuntimeMins `
                -MaximumRuntimeMins $MaximumRuntimeMins `
                -ContentFallback `
                -SlowNetworkDeploymentMode Download `
                -ErrorAction Stop | Out-Null

            Write-Log "  Created: Override=$($entry.OverrideHex) HyperV=$($entry.IsHyperV) Requirements=$($requirements.Count) DetectionClauses=$($clauses.Count)"
        }

        # --- Verify server-side: every matrix entry must exist as a deployment type ---
        $step = "Deployment type verification ('$AppName')"
        $serverDts = @(Get-CMDeploymentType -ApplicationName $AppName -ErrorAction SilentlyContinue)
        Write-Log ""
        Write-Log "Deployment types on site     : $($serverDts.Count) of $($matrix.Count) expected"
        $missingDts = @()
        foreach ($entry in $matrix) {
            if (@($serverDts | Where-Object { $_.LocalizedDisplayName -eq $entry.Name }).Count -gt 0) {
                Write-Log "  OK      $($entry.Name)"
            }
            else {
                $missingDts += $entry.Name
                Write-Log "  MISSING $($entry.Name)" -Level ERROR
            }
        }
        if ($missingDts.Count -gt 0) {
            throw "Deployment type verification failed; missing: $($missingDts -join '; ')"
        }

        $step = "Remove-CMApplicationRevisionHistory (CI_ID=$($cmApp.CI_ID))"
        Remove-CMApplicationRevisionHistoryByCIId -CI_ID ([UInt32]$cmApp.CI_ID) -KeepLatest 1

        Write-Log ""
        Write-Log "Created MECM application     : $AppName ($($matrix.Count) deployment types)"
        Write-Log ""
        Write-Log "Next:" -Level WARN
        Write-Log "  1. Distribute content to your DP group."
        Write-Log "  2. Deploy the application (Required) to the SpecExec collections or a broad"
        Write-Log "     limiting collection; requirement rules route each device to its deployment type."
        Write-Log "  3. Values apply at next reboot (exit 3010 / BasedOnExitCode)."

        return [UInt32]$cmApp.CI_ID
    }
    catch {
        Write-LogErrorRecord -ErrorRecord $_ -Context ("Invoke-PackageSpecExec failed during step: {0}" -f $step)
        throw
    }
}

# --- Latest-only mode ---
# No vendor download; the "version" is the pinned mitigation content version.
if ($GetLatestVersionOnly) {
    Write-Output $ContentVersion
    exit 0
}

# --- Main ---
try {
    $startLocation = Get-Location

    Write-Log ""
    Write-Log ("=" * 60)
    Write-Log "SpecExec Mitigations Packager starting"
    Write-Log ("=" * 60)
    Write-Log ""
    Write-Log ("RunAsUser                    : {0}\{1}" -f $env:USERDOMAIN, $env:USERNAME)
    Write-Log ("Machine                      : {0}" -f $env:COMPUTERNAME)
    Write-Log "Start location               : $startLocation"
    Write-Log "SiteCode                     : $SiteCode"
    Write-Log "FileServerPath               : $FileServerPath"
    Write-Log "BaseDownloadRoot             : $BaseDownloadRoot"
    Write-Log "ContentVersion               : $ContentVersion"
    Write-Log ""

    if ($StageOnly) {
        Invoke-StageSpecExec
    }
    elseif ($PackageOnly) {
        Invoke-PackageSpecExec
    }
    else {
        Invoke-StageSpecExec
        Invoke-PackageSpecExec
    }

    Write-Log ""
    Write-Log "Script execution complete."
}
catch {
    Write-LogErrorRecord -ErrorRecord $_ -Context 'package-specexec-mitigations'
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -Level ERROR
    exit 1
}
finally {
    Set-Location $startLocation -ErrorAction SilentlyContinue
}
