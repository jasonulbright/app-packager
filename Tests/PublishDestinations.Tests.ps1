#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Every deployment target publishes to exactly its destinations and never touches the others.

.DESCRIPTION
    ConfigMgr, Intune and WSUS are independent and optional. The three publishers are mocked: a run connects
    to ConfigMgr only when the target names ConfigMgr, calls Graph only when it names Intune and calls WSUS
    only when it names WSUS. A target without ConfigMgr runs with no site code, no content share and no
    ConfigMgr module. A value that is no target name selects no destination.
#>

BeforeDiscovery {
    $script:Combinations = @(
        @{ Target = 'ConfigMgr';                 Cm = $true;  Intune = $false; Wsus = $false }
        @{ Target = 'IntuneOnly';                Cm = $false; Intune = $true;  Wsus = $false }
        @{ Target = 'WSUSOnly';                  Cm = $false; Intune = $false; Wsus = $true }
        @{ Target = 'ConfigMgrAndIntune';        Cm = $true;  Intune = $true;  Wsus = $false }
        @{ Target = 'ConfigMgrAndWSUS';          Cm = $true;  Intune = $false; Wsus = $true }
        @{ Target = 'IntuneAndWSUS';             Cm = $false; Intune = $true;  Wsus = $true }
        @{ Target = 'ConfigMgrAndIntuneAndWSUS'; Cm = $true;  Intune = $true;  Wsus = $true }
    )
}

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerOneClick.psd1') -Force -Global
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$null, [ref]$null)
    foreach ($name in 'Invoke-PackagerPackage', 'Invoke-PackagerPackageWithConflictPrompt', 'Test-DeploymentTargetSkipsSite', 'Test-DeploymentTargetPublishesToWsus',
        'Test-DeploymentTargetPublishesToIntune', 'Get-DeploymentTargetNames', 'Get-KnownDeploymentTarget', 'ConvertTo-DeploymentTargetName', 'Get-PublishRowOutcome') {
        $node = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $node) { throw "start-apppackager.ps1 no longer defines $name." }
        # The text box of the window is not needed here and would need the WPF assemblies.
        . ([scriptblock]::Create($node.Extent.Text.Replace('[System.Windows.Controls.TextBox]$LogTextBox = $null', '$LogTextBox = $null')))
    }

    # The child process, the share check and the three publishers are stand-ins.
    function Set-ProcessStartInfoArgumentList { param($StartInfo, $Arguments) $script:ChildArguments = @($Arguments) }
    function Set-PackagerEnvironment { param([Parameter(ValueFromRemainingArguments)]$Rest) }
    function Test-PackagerSupportsFileServerPath { param($PackagerPath) $true }
    function Invoke-ProcessWithStreaming { param($StartInfo, $OutLog, $ErrLog, $StructuredLog, $LogTextBox) [pscustomobject]@{ ExitCode = 0; StdOut = 'ok' } }
    function Assert-PackagerPackageIntegrity { param([Parameter(ValueFromRemainingArguments)]$Rest) }
    function Invoke-PackagerIntuneWinPostStep { param([Parameter(ValueFromRemainingArguments)]$Rest) [pscustomobject]@{ Ok = $true; Message = 'built' } }
    function Invoke-PackagerIntunePublish { param([Parameter(ValueFromRemainingArguments)]$Rest) [pscustomobject]@{ Ok = $true } }
    function Invoke-PackagerWsusPostStep { param([Parameter(ValueFromRemainingArguments)]$Rest) [pscustomobject]@{ Ok = $true; Refused = $false; Warnings = @() } }
    function Get-MecmCurrentVersionByCMName { param([Parameter(ValueFromRemainingArguments)]$Rest) [pscustomobject]@{ Found = $false } }
    function Connect-CMSite { param([Parameter(ValueFromRemainingArguments)]$Rest) $true }
}

Describe 'A package run reaches exactly the destinations of its target' {
    BeforeEach {
        Mock Invoke-PackagerIntuneWinPostStep { [pscustomobject]@{ Ok = $true; Message = 'built' } }
        Mock Invoke-PackagerIntunePublish { [pscustomobject]@{ Ok = $true } }
        Mock Invoke-PackagerWsusPostStep { [pscustomobject]@{ Ok = $true; Refused = $false; Warnings = @() } }
        Mock Assert-PackagerPackageIntegrity { }
        $script:ChildArguments = @()
    }

    It 'runs <Target> with ConfigMgr: <Cm>, Intune: <Intune>, WSUS: <Wsus>' -TestCases $script:Combinations {
        param($Target, $Cm, $Intune, $Wsus)
        $packageArgs = @{
            PackagerPath = 'C:\p\package-x.ps1'; LogFolder = $TestDrive; DownloadRoot = $TestDrive
            SiteCode = $(if ($Cm) { 'XYZ' } else { '' }); FileServerPath = $(if ($Cm) { '\\server\share' } else { '' })
            DeploymentTarget = $Target
            IntunePublishConfig = $(if ($Intune) { @{ TenantId = 't'; ClientId = 'c'; ClientSecret = 's' } } else { $null })
            WsusPublishConfig = $(if ($Wsus) { @{ ServerName = 'wsus01' } } else { $null })
        }
        $result = Invoke-PackagerPackage @packageArgs
        $result.ExitCode | Should -Be 0

        # ConfigMgr: the package phase connects to the site; any other run stops at Stage.
        if ($Cm) { $script:ChildArguments | Should -Contain '-PackageOnly'; $script:ChildArguments | Should -Contain '-SiteCode' }
        else { $script:ChildArguments | Should -Contain '-StageOnly'; $script:ChildArguments | Should -Not -Contain '-SiteCode'; $script:ChildArguments | Should -Not -Contain '-PackageOnly' }
        Should -Invoke Assert-PackagerPackageIntegrity -Times ([int]$Cm) -Exactly
        # Intune: the Graph publish. WSUS: the WSUS publish.
        Should -Invoke Invoke-PackagerIntunePublish -Times ([int]$Intune) -Exactly
        Should -Invoke Invoke-PackagerWsusPostStep -Times ([int]$Wsus) -Exactly
        if ($Intune) { $result.IntunePublish.Ok | Should -BeTrue } else { $result.PSObject.Properties['IntunePublish'] | Should -BeNullOrEmpty }
        if ($Wsus) { $result.WsusPublish.Ok | Should -BeTrue } else { $result.PSObject.Properties['WsusPublish'] | Should -BeNullOrEmpty }
    }

    It 'runs the old name <Old> as <Current>' -TestCases @(
        @{ Old = 'MECM'; Current = 'ConfigMgr'; Wsus = 0 }
        @{ Old = 'MECMAndWSUS'; Current = 'ConfigMgrAndWSUS'; Wsus = 1 }
        @{ Old = 'MECMAndIntune'; Current = 'ConfigMgrAndIntune'; Wsus = 0 }
    ) {
        param($Old, $Current, $Wsus)
        $null = Invoke-PackagerPackage -PackagerPath 'C:\p\package-x.ps1' -LogFolder $TestDrive -DownloadRoot $TestDrive -SiteCode 'XYZ' -FileServerPath '\\s\s' -DeploymentTarget $Old `
            -IntunePublishConfig @{ TenantId = 't' } -WsusPublishConfig @{ ServerName = 'w' }
        $script:ChildArguments | Should -Contain '-PackageOnly'
        Should -Invoke Invoke-PackagerWsusPostStep -Times $Wsus -Exactly
    }

    It 'asks for no site and no ConfigMgr query before a <Target> run' -TestCases @(
        $script:Combinations | Where-Object { -not $_.Cm }
    ) {
        param($Target)
        Mock Get-MecmCurrentVersionByCMName { throw 'must not query the site' }
        Mock Connect-CMSite { throw 'must not connect to the site' }
        Mock Invoke-PackagerPackage { [pscustomobject]@{ ExitCode = 0; StdOut = 'Packaged' } }
        $state = @{ ConflictDecisionForAll = 'Skip'; CancelRequested = $false; LogQueue = (New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]') }
        $null = Invoke-PackagerPackageWithConflictPrompt -State $state -AppLabel Browser -PackageArgs @{ PackagerPath = 'C:\p\package-x.ps1'; LogFolder = $TestDrive; SiteCode = ''; FileServerPath = ''; DeploymentTarget = $Target }
        Should -Invoke Get-MecmCurrentVersionByCMName -Times 0 -Exactly
        Should -Invoke Connect-CMSite -Times 0 -Exactly
        Should -Invoke Invoke-PackagerPackage -Times 1 -Exactly -ParameterFilter { -not $Preflight }
    }
}

Describe 'A value that is no target name selects no destination' {
    It 'selects no destination for <Value>' -TestCases @(
        @{ Value = '' }
        @{ Value = '   ' }
        @{ Value = 'Nonsense' }
        @{ Value = 'Intune' }
        @{ Value = 'MECMAndIntuneAndWSUS' }
    ) {
        param($Value)
        Get-DeploymentTargetDestinations -Target $Value | Should -BeNullOrEmpty
        Test-DeploymentTargetPublishesToWsus -DeploymentTarget $Value | Should -BeFalse
        Test-DeploymentTargetPublishesToIntune -DeploymentTarget $Value | Should -BeFalse
        # No ConfigMgr destination: the value never reaches the site.
        Test-DeploymentTargetSkipsSite -DeploymentTarget $Value | Should -BeTrue
    }

    It 'stops a run with the valid names for <Value>' -TestCases @(@{ Value = '' }, @{ Value = 'Nonsense' }) {
        param($Value)
        { Get-KnownDeploymentTarget -Value $Value } | Should -Throw '*ConfigMgr, ConfigMgrAndIntune, IntuneOnly, ConfigMgrAndWSUS, WSUSOnly, IntuneAndWSUS, ConfigMgrAndIntuneAndWSUS*'
    }

    It 'does not map an unknown name to ConfigMgr in the One Click set' {
        $set = ConvertTo-OneClickDestinationSet 'Nonsense'
        ($set.ConfigMgr -or $set.WSUS -or $set.Intune) | Should -BeFalse
    }

    It 'does not map an unknown name to ConfigMgr in the run snapshot target' {
        ConvertTo-CurrentDeploymentTarget -Value 'Nonsense' -Default 'ContentOnly' | Should -BeExactly 'ContentOnly'
    }
}

Describe 'The name of a set of destinations' {
    It 'names ConfigMgr: <Cm>, Intune: <Intune>, WSUS: <Wsus> as <Name>' -TestCases @(
        @{ Cm = $true;  Intune = $false; Wsus = $false; Name = 'ConfigMgr' }
        @{ Cm = $false; Intune = $true;  Wsus = $false; Name = 'IntuneOnly' }
        @{ Cm = $false; Intune = $false; Wsus = $true;  Name = 'WSUSOnly' }
        @{ Cm = $true;  Intune = $true;  Wsus = $false; Name = 'ConfigMgrAndIntune' }
        @{ Cm = $true;  Intune = $false; Wsus = $true;  Name = 'ConfigMgrAndWSUS' }
        @{ Cm = $false; Intune = $true;  Wsus = $true;  Name = 'IntuneAndWSUS' }
        @{ Cm = $true;  Intune = $true;  Wsus = $true;  Name = 'ConfigMgrAndIntuneAndWSUS' }
        @{ Cm = $false; Intune = $false; Wsus = $false; Name = '' }
    ) {
        param($Cm, $Intune, $Wsus, $Name)
        $set = [pscustomobject]@{ ConfigMgr = $Cm; Intune = $Intune; WSUS = $Wsus }
        ConvertTo-DeploymentTargetName -Destinations $set | Should -BeExactly $Name
        if ($Name) {
            $back = Get-DeploymentTargetDestinations -Target $Name
            $back.ConfigMgr | Should -Be $Cm
            $back.Intune | Should -Be $Intune
            $back.WSUS | Should -Be $Wsus
            $one = ConvertTo-OneClickDestinationSet $Name
            $one.ConfigMgr | Should -Be $Cm
            $one.Intune | Should -Be $Intune
            $one.WSUS | Should -Be $Wsus
        }
    }

    It 'lists the seven publish names' {
        @(Get-DeploymentTargetNames).Count | Should -Be 7
        @(Get-DeploymentTargetNames) | Should -Contain 'IntuneAndWSUS'
        @(Get-DeploymentTargetNames) | Should -Contain 'ConfigMgrAndIntuneAndWSUS'
    }
}

Describe 'The row outcome of a run without ConfigMgr' {
    It 'reports an Intune and WSUS run as published when both publishes succeed' {
        $result = [pscustomobject]@{ ExitCode = 0; IntunePublish = [pscustomobject]@{ Ok = $true }; WsusPublish = [pscustomobject]@{ Ok = $true; Refused = $false; Warnings = @() } }
        $outcome = Get-PublishRowOutcome -Result $result -DeploymentTarget 'IntuneAndWSUS'
        $outcome.Kind | Should -Be 'Success'
        $outcome.Published | Should -BeTrue
    }

    It 'reports an Intune and WSUS run as failed when the Intune publish fails' {
        $result = [pscustomobject]@{ ExitCode = 0; IntunePublish = [pscustomobject]@{ Ok = $false }; WsusPublish = [pscustomobject]@{ Ok = $true; Refused = $false; Warnings = @() } }
        (Get-PublishRowOutcome -Result $result -DeploymentTarget 'IntuneAndWSUS').Kind | Should -Be 'Failed'
    }

    It 'reports a three-destination run as packaged with the Intune failure noted' {
        $result = [pscustomobject]@{ ExitCode = 0; WsusPublish = [pscustomobject]@{ Ok = $true; Refused = $false; Warnings = @() } }
        $outcome = Get-PublishRowOutcome -Result $result -DeploymentTarget 'ConfigMgrAndIntuneAndWSUS'
        $outcome.Kind | Should -Be 'Success'
        $outcome.Status | Should -Match 'Intune failed'
    }
}
