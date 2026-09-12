#Requires -Version 5.1
# The packager child bridge: the selected build has to be the content the
# child would resolve, and an unattended batch run has to carry the signing
# policy the same way an interactive one does. A fixture packager stands in
# for a catalog script; nothing contacts a site.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Import-Module (Join-Path $root 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking
Import-Module (Join-Path $root 'Packagers\AppPackagerCommon.psm1') -Force -DisableNameChecking
Import-Module (Join-Path $root 'Packagers\AppPackagerWorkbench.psm1') -Force -DisableNameChecking

$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
if ($errors) { throw ($errors.Message -join '; ') }
$published = New-Object System.Collections.Generic.List[string]
foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
    Set-Item -Path ('function:global:' + $fn.Name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
    $published.Add($fn.Name)
}

$sandbox = Join-Path $env:TEMP ('ap-bridge-' + [guid]::NewGuid().ToString('N'))
$dataRoot = Join-Path $sandbox 'data'
$downloadRoot = Join-Path $sandbox 'dl'
$logFolder = Join-Path $sandbox 'logs'
$fixtureRoot = Join-Path $sandbox 'Packagers'
$env:APP_PACKAGER_WORKBENCH_ROOT = $dataRoot
foreach ($d in @($dataRoot, $downloadRoot, $logFolder, $fixtureRoot)) { [void](New-Item -ItemType Directory -Path $d -Force) }

$checks = 0
$assert = {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw ('FAIL: ' + $Message) }
    $script:checks += 1
    Write-Host ('  ok  ' + $Message)
}

try {
    $script:Prefs = [pscustomobject]@{
        EstimatedRuntimeMins = 15
        MaximumRuntimeMins   = 30
        DownloadRoot         = $downloadRoot
        ScriptSigning        = [pscustomobject]@{
            SignDetection = $true; SignRequirements = $false; SignDeployment = $false
            RequireDetection = $true; RequireRequirements = $false; RequireDeployment = $false
            CertificateThumbprint = ''; StoreLocation = 'CurrentUser'
            TimestampServer = ''; TimestampRequired = $false; HashAlgorithm = 'SHA256'
        }
    }
    function Write-Log { param([string]$Message, [string]$Level) $script:BatchLog += $Message }
    function Add-LogLine { param([string]$Message) }
    $script:BatchLog = @()

    # A fixture packager: it echoes the bridge variables and exits. The
    # BaseDownloadRoot line is what Get-PackagerFolderInfo reads.
    $fixturePath = Join-Path $fixtureRoot 'package-bridgefixture.ps1'
    @(
        'param([switch]$StageOnly, [switch]$PackageOnly, [string]$SiteCode, [string]$Comment,'
        '      [string]$LogPath, [string]$DownloadRoot, [string]$ContentLayout, [string]$FileServerPath)'
        '# Vendor: Fixture'
        '# Application: Bridge Fixture'
        '$BaseDownloadRoot = Join-Path $DownloadRoot "Fixture\BridgeFixture"'
        'if (-not $StageOnly -and -not $PackageOnly) { Write-Output "2.0.0"; exit 0 }'
        'Write-Output ("SIGNING=" + $env:APP_PACKAGER_SIGNING)'
        'Write-Output ("SNAPSHOT=" + $env:APP_PACKAGER_RUN_SNAPSHOT)'
        'Write-Output ("WBROOT=" + $env:APP_PACKAGER_WORKBENCH_ROOT)'
        'Write-Output ("DLROOT=" + $env:APP_PACKAGER_DOWNLOAD_ROOT)'
        'Write-Output ("BUILDID=" + $env:APP_PACKAGER_BUILD_ID)'
        'exit 0'
    ) | Set-Content -LiteralPath $fixturePath -Encoding ASCII

    $stageRoot = Join-Path $downloadRoot 'Fixture\BridgeFixture'
    $writeStage = {
        param([string]$Version, [string]$BuildId)
        $folder = Join-Path $stageRoot $Version
        [void](New-Item -ItemType Directory -Path $folder -Force)
        @{ SchemaVersion = 4; BuildId = $BuildId; SoftwareVersion = $Version } |
            ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $folder 'stage-manifest.json') -Encoding UTF8
        Start-Sleep -Milliseconds 1100
        return (Join-Path $folder 'stage-manifest.json')
    }

    # --- Build selection -------------------------------------------------
    $oldManifest = & $writeStage '1.0.0' '20260101-000000-aaaaaaaa'
    $newManifest = & $writeStage '2.0.0' '20260202-000000-bbbbbbbb'

    $selected = Assert-WorkbenchBuildSelection -BuildId '20260202-000000-bbbbbbbb' -DownloadRoot $downloadRoot -PackagerPath $fixturePath
    & $assert ([string]$selected.Path -eq $newManifest) 'the newest build resolves to its own manifest'

    $refused = $null
    try { [void](Assert-WorkbenchBuildSelection -BuildId '20260101-000000-aaaaaaaa' -DownloadRoot $downloadRoot -PackagerPath $fixturePath) }
    catch { $refused = $_.Exception.Message }
    & $assert ($refused -like 'stale build:*') 'selecting a build the child would not resolve is refused as stale'
    & $assert ($refused -like '*20260202-000000-bbbbbbbb*') 'the refusal names the build the child would package instead'

    $missing = $null
    try { [void](Assert-WorkbenchBuildSelection -BuildId 'never-staged' -DownloadRoot $downloadRoot -PackagerPath $fixturePath) }
    catch { $missing = $_.Exception.Message }
    & $assert ($missing -like '*stale build*') 'a build id with no staged manifest is refused'

    $scoped = Get-WorkbenchPackagerStageRoot -DownloadRoot $downloadRoot -PackagerPath $fixturePath
    & $assert ($scoped -eq $stageRoot) 'the search is scoped to the packager subtree the child reads'

    # The selected build reaches the child as the two new bridge variables.
    $result = Invoke-PackagerPackage -PackagerPath $fixturePath -SiteCode 'ZZZ' -FileServerPath $sandbox `
        -LogFolder $logFolder -DownloadRoot $downloadRoot -BuildId '20260202-000000-bbbbbbbb' `
        -SigningJson (Get-WorkbenchSigningPolicyJson) -WorkbenchDataRoot $dataRoot -Preflight
    & $assert ($result.StdOut -match 'BUILDID=20260202-000000-bbbbbbbb') 'the child receives APP_PACKAGER_BUILD_ID'
    & $assert ($result.StdOut -match 'DLROOT=') 'the child receives APP_PACKAGER_DOWNLOAD_ROOT'
    & $assert ($result.StdOut -match 'SIGNING=.*SignDetection') 'the child receives the signing policy'

    $profileRoot = Get-WorkbenchProfileDownloadRoot -DownloadRoot $downloadRoot -ProfileId 'abc123'
    & $assert ($profileRoot -eq (Join-Path (Join-Path $downloadRoot 'profiles') 'abc123')) 'a non-default profile packages from its own staging root'
    & $assert ((Get-WorkbenchProfileDownloadRoot -DownloadRoot $downloadRoot -ProfileId 'default') -eq $downloadRoot) 'the default profile keeps the existing staging root'

    # --- Unattended batch carries the same bridge -------------------------
    # A require switch with no usable certificate stops the batch before the
    # child starts, so the refusal is asserted first and the bridge itself is
    # proven with every switch off.
    $strictJson = Get-WorkbenchSigningPolicyJson
    $script:BatchLog = @()
    [void](Invoke-BatchUpdate -PackagersRoot $fixtureRoot -Apps @('package-bridgefixture') `
        -OnUpdateFound 'Stage' -SiteCode 'ZZZ' -FileServerPath $sandbox -DownloadRoot $downloadRoot `
        -SigningJson $strictJson -Force)
    $strictText = ($script:BatchLog -join "`n")
    & $assert ($strictText -match 'signing policy cannot be met') 'a require switch without a certificate is refused before the batch child runs'
    & $assert ($strictText -notmatch 'SIGNING=') 'the refused application never reaches the child'

    $script:Prefs.ScriptSigning.SignDetection = $false
    $script:Prefs.ScriptSigning.RequireDetection = $false
    $signingJson = Get-WorkbenchSigningPolicyJson
    $script:BatchLog = @()
    $beforeSigning = $env:APP_PACKAGER_SIGNING
    [void](Invoke-BatchUpdate -PackagersRoot $fixtureRoot -Apps @('package-bridgefixture') `
        -OnUpdateFound 'Stage' -SiteCode 'ZZZ' -FileServerPath $sandbox -DownloadRoot $downloadRoot `
        -SigningJson $signingJson -Force)
    $batchText = ($script:BatchLog -join "`n")
    & $assert ($batchText -match "SIGNING=.*SignDetection") "the batch child receives the signing policy"
    & $assert ($batchText -match 'WBROOT=') 'the batch child receives the workbench data root'
    & $assert ($batchText -match 'DLROOT=') 'the batch child receives the download root'
    & $assert ($batchText -match 'SNAPSHOT=\S') 'the batch child receives a run snapshot for its active profile'
    & $assert ($env:APP_PACKAGER_SIGNING -eq $beforeSigning) 'the batch run restores the process environment'

    # --- The drop path leaves no policy on the GUI process ----------------
    $adHocSource = ($ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-DropIntake' }, $false)).Extent.Text
    & $assert ($adHocSource -notmatch '\$env:APP_PACKAGER_SIGNING\s*=') 'the drop path no longer assigns the policy on the GUI process'
    & $assert ($adHocSource -notmatch '\$env:APP_PACKAGER_WORKBENCH_ROOT\s*=') 'the drop path no longer assigns the data root on the GUI process'
    $adHocPipeline = ($ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-AdHocPipeline' }, $false)).Extent.Text
    & $assert ($adHocPipeline -match 'savedSigning' -and $adHocPipeline -match 'Remove-Item Env:\\APP_PACKAGER_SIGNING') 'the in-process ad-hoc run restores the policy in finally'

    # --- Discovery and migration agree on both discovery roots ------------
    $workbenchSource = ($ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Show-ApplicationWorkbench' }, $false)).Extent.Text
    & $assert ($workbenchSource -match 'Invoke-LegacyPreferenceMigration[^\r\n]*-PackagersRoot[^\r\n]*-CustomScriptRoot') 'migration receives both discovery roots'
    $listSource = ($ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-WorkbenchUiApplications' }, $false)).Extent.Text
    & $assert ($listSource -match 'Get-WorkbenchApplications[^\r\n]*-PackagersRoot[^\r\n]*-CustomScriptRoot') 'discovery receives the same two roots'
    & $assert ((Get-WorkbenchCustomScriptRoot) -eq (Join-Path $dataRoot 'scripts')) 'the custom script root resolves under the workbench data root'

    Write-Host ''
    Write-Host ('PASS: packager child bridge, {0} checks' -f $script:checks)
}
finally {
    foreach ($name in @($published)) { Remove-Item -Path ('function:global:' + $name) -Force -ErrorAction SilentlyContinue }
    Remove-Item -LiteralPath $sandbox -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item Env:\APP_PACKAGER_WORKBENCH_ROOT -ErrorAction SilentlyContinue
}
