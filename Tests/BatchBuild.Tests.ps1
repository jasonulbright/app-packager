#Requires -Modules Pester

<#
.SYNOPSIS
    Headless batch builds: what counts as already built, and where a named
    profile stages.

.DESCRIPTION
    Invoke-BatchUpdate runs out of start-apppackager.ps1, so its functions are
    lifted out by AST and published for the run. A fixture packager stands in
    for a catalog script: it reports a version from the environment, records
    the -DownloadRoot it was given, and can be told to fail. History lives
    under a scratch LOCALAPPDATA.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    Import-Module (Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psd1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerWorkbench.psd1') -Force -DisableNameChecking

    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors) { throw ($errors.Message -join '; ') }
    $script:Published = New-Object System.Collections.Generic.List[string]
    foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)) {
        Set-Item -Path ('function:global:' + $fn.Name) -Value ([scriptblock]::Create($fn.Body.Extent.Text.Trim('{', '}')))
        $script:Published.Add($fn.Name)
    }
    function global:Write-Log { param([string]$Message, [string]$Level) $global:BatchBuildLog += $Message }
    $global:BatchBuildLog = @()

    $script:Sandbox = Join-Path $env:TEMP ('ap-batch-' + [guid]::NewGuid().ToString('N'))
    $script:DataRoot = Join-Path $script:Sandbox 'data'
    $script:DownloadRoot = Join-Path $script:Sandbox 'dl'
    $script:PackagersRoot = Join-Path $script:Sandbox 'Packagers'
    $script:SavedLocalAppData = $env:LOCALAPPDATA
    $script:SavedWorkbenchRoot = $env:APP_PACKAGER_WORKBENCH_ROOT
    foreach ($d in @($script:DataRoot, $script:DownloadRoot, $script:PackagersRoot, (Join-Path $script:Sandbox 'local'))) { [void](New-Item -ItemType Directory -Path $d -Force) }
    $env:LOCALAPPDATA = Join-Path $script:Sandbox 'local'
    $env:APP_PACKAGER_WORKBENCH_ROOT = $script:DataRoot

    $script:CallLog = Join-Path $script:Sandbox 'calls.log'
    $fixture = @'
param([switch]$StageOnly, [switch]$PackageOnly, [switch]$GetLatestVersionOnly, [string]$SiteCode, [string]$Comment,
      [string]$LogPath, [string]$DownloadRoot, [string]$ContentLayout, [string]$FileServerPath,
      [int]$EstimatedRuntimeMins, [int]$MaximumRuntimeMins)
# Vendor: Fixture
# Application: Batch Fixture
$BaseDownloadRoot = Join-Path $DownloadRoot "Fixture\BatchFixture"
if ($GetLatestVersionOnly) { Write-Output $env:APP_TEST_VERSION; exit 0 }
Add-Content -LiteralPath $env:APP_TEST_CALLS -Value ("ROOT=" + $DownloadRoot + ";VERSION=" + $env:APP_TEST_VERSION)
if ($env:APP_TEST_FAIL -eq '1') { Write-Error 'fixture stage failed'; exit 1 }
exit 0
'@
    [System.IO.File]::WriteAllText((Join-Path $script:PackagersRoot 'package-batchfixture.ps1'), $fixture, (New-Object System.Text.ASCIIEncoding))
    $env:APP_TEST_CALLS = $script:CallLog

    $script:RunBatch = {
        param([string]$Version, [string]$Action = 'Stage', [switch]$Fail, [switch]$Force)
        $env:APP_TEST_VERSION = $Version
        $env:APP_TEST_FAIL = $(if ($Fail) { '1' } else { '0' })
        $global:BatchBuildLog = @()
        $results = Invoke-BatchUpdate -PackagersRoot $script:PackagersRoot -Apps @('package-batchfixture') `
            -OnUpdateFound $Action -SiteCode 'ZZZ' -FileServerPath $script:Sandbox -DownloadRoot $script:DownloadRoot -Force:$Force
        return @($results)[0]
    }
    $script:CallCount = { if (Test-Path -LiteralPath $script:CallLog) { @(Get-Content -LiteralPath $script:CallLog).Count } else { 0 } }
}

AfterAll {
    foreach ($name in @($script:Published)) { Remove-Item -Path ('function:global:' + $name) -Force -ErrorAction SilentlyContinue }
    Remove-Item -Path 'function:global:Write-Log' -Force -ErrorAction SilentlyContinue
    $env:LOCALAPPDATA = $script:SavedLocalAppData
    $env:APP_PACKAGER_WORKBENCH_ROOT = $script:SavedWorkbenchRoot
    foreach ($name in 'APP_TEST_CALLS', 'APP_TEST_VERSION', 'APP_TEST_FAIL') { Remove-Item -Path ('Env:\' + $name) -ErrorAction SilentlyContinue }
    Remove-Item -LiteralPath $script:Sandbox -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'Batch gate' {
    It 'stages a version that a check saw after the previous version was built' {
        $first = & $script:RunBatch '1.0.0'
        $first.Action | Should -Be 'Stage'
        (& $script:CallCount) | Should -Be 1

        $seen = & $script:RunBatch '2.0.0' 'Report' -Force
        $seen.Action | Should -Be 'Reported'
        (& $script:CallCount) | Should -Be 1

        $next = & $script:RunBatch '2.0.0'
        $next.Action | Should -Be 'Stage'
        (& $script:CallCount) | Should -Be 2
    }

    It 'skips a version that was staged' {
        $again = & $script:RunBatch '2.0.0'
        $again.Action | Should -Be 'NoChange'
        (& $script:CallCount) | Should -Be 2
    }

    It 'retries a stage that failed' {
        $failed = & $script:RunBatch '3.0.0' 'Stage' -Fail
        $failed.Action | Should -Be 'Failed'
        (& $script:CallCount) | Should -Be 3

        $retry = & $script:RunBatch '3.0.0'
        $retry.Action | Should -Be 'Stage'
        (& $script:CallCount) | Should -Be 4

        (& $script:RunBatch '3.0.0').Action | Should -Be 'NoChange'
    }

    It 'stages again when the version was only packaged for another destination' {
        $packaged = & $script:RunBatch '3.0.0' 'StageAndPackage'
        $packaged.Action | Should -Be 'StageAndPackage'
        (& $script:CallCount) | Should -Be 5
        (& $script:RunBatch '3.0.0' 'StageAndPackage').Action | Should -Be 'NoChange'
    }

    It 'rebuilds a version whose build record belongs to another profile revision or policy' {
        $original = Get-Command Get-LatestBuildRecord -CommandType Function
        try {
            (& $script:RunBatch '6.0.0').Action | Should -Be 'Stage'
            (& $script:RunBatch '6.0.0').Action | Should -Be 'NoChange'

            Set-Item -Path 'function:global:Get-LatestBuildRecord' -Value { [pscustomobject]@{ SoftwareVersion = '6.0.0'; ProfileRevision = 7; PolicyDigest = '' } }
            (& $script:RunBatch '6.0.0').Action | Should -Be 'Stage'

            Set-Item -Path 'function:global:Get-LatestBuildRecord' -Value { [pscustomobject]@{ SoftwareVersion = '6.0.0'; ProfileRevision = 0; PolicyDigest = '' } }
            (& $script:RunBatch '6.0.0').Action | Should -Be 'NoChange'
        }
        finally {
            Remove-Item -Path 'function:global:Get-LatestBuildRecord' -ErrorAction SilentlyContinue
            if ($original -and $original.Module) { Import-Module $original.Module -Force -Global -DisableNameChecking }
        }
    }
    It 'treats a legacy entry whose last check is newer than its stage as not built' {
        $entry = [pscustomobject]@{ LastChecked = '2026-10-02T00:00:00Z'; LastStaged = '2026-10-01T00:00:00Z'; LastPackaged = $null; LastKnownVersion = '2.0' }
        Get-PackagerHistoryActedVersion -Entry $entry -Event Staged | Should -Be ''
        $entry.LastStaged = '2026-10-02T00:00:00Z'
        Get-PackagerHistoryActedVersion -Entry $entry -Event Staged | Should -Be '2.0'
    }

    It 'treats a version recorded for another scope as not built' {
        $entry = @{ LastStagedVersion = '5.0'; LastStagedScope = 'D:\one'; LastKnownVersion = '5.0' }
        Get-PackagerHistoryActedVersion -Entry $entry -Event Staged -Scope 'D:\one' | Should -Be '5.0'
        Get-PackagerHistoryActedVersion -Entry $entry -Event Staged -Scope 'D:\two' | Should -Be ''
    }
}

Describe 'Batch profile staging root' {
    It 'launches the packager with the named profile root' {
        $appId = New-ApplicationId -Kind Catalog -ScriptPath (Join-Path $script:PackagersRoot 'package-batchfixture.ps1')
        $profile = New-WorkbenchProfileObject -ApplicationId $appId -Name 'Isolated'
        $saved = Save-Profile -Profile ([pscustomobject]$profile) -DataRoot $script:DataRoot -SetActive

        Remove-Item -LiteralPath $script:CallLog -Force -ErrorAction SilentlyContinue
        $r = & $script:RunBatch '9.0.0' 'Stage' -Fail:$false
        $r.Action | Should -Be 'Stage'

        $expectedRoot = Join-Path (Join-Path $script:DownloadRoot 'profiles') $saved.ProfileId
        (Get-Content -LiteralPath $script:CallLog -Raw) | Should -Match ([regex]::Escape('ROOT=' + $expectedRoot))
    }

    It 'does not treat the default profile build as the named profile build' {
        $entryAfterNamed = (Read-PackagerHistory)['package-batchfixture']
        $entryAfterNamed.LastStagedScope | Should -Match 'profiles'
        Get-PackagerHistoryActedVersion -Entry $entryAfterNamed -Event Staged -Scope $script:DownloadRoot | Should -Be ''
    }
}
