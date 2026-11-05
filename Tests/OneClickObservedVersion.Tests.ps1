#Requires -Modules Pester

<#
.SYNOPSIS
    One Click checks the version the site holds, and the site query uses the
    provider that was asked for.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $tokens = $null; $errors = $null
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors) { throw ($errors.Message -join '; ') }
    foreach ($name in 'Compare-SemVer', 'Get-OneClickObservedSkipReason', 'Test-MecmDriveRoot', 'Get-MecmCurrentVersionByCMName', 'Get-MecmApplicationNamePattern') {
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }.GetNewClosure(), $false)
        if ($fn) { . ([scriptblock]::Create($fn.Extent.Text)) }
        elseif ($name -ne 'Get-MecmApplicationNamePattern') { throw "function $name not found" }
    }
}

Describe 'The site version decides a One Click skip, not the publish record alone' {
    It 'skips ConfigMgr when the site holds a newer version than the vendor reports' {
        Get-OneClickObservedSkipReason -Destination 'ConfigMgr' -Observed '3.0' -Latest '2.0' | Should -Be 'ConfigMgr: site holds 3.0, newer than 2.0'
    }
    It 'does not skip an older or equal site version, an unknown version, or a forced run' {
        Get-OneClickObservedSkipReason -Destination 'ConfigMgr' -Observed '1.0' -Latest '2.0' | Should -Be ''
        Get-OneClickObservedSkipReason -Destination 'ConfigMgr' -Observed '2.0' -Latest '2.0' | Should -Be ''
        Get-OneClickObservedSkipReason -Destination 'ConfigMgr' -Observed '' -Latest '2.0' | Should -Be ''
        Get-OneClickObservedSkipReason -Destination 'ConfigMgr' -Observed '3.0' -Latest '2.0' -Force | Should -Be ''
    }
    It 'reads the site in the One Click branch and applies the reason in the destination loop' {
        $branch = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.SwitchStatementAst] -and $n.Extent.Text -match "'OneClick'" -and $n.Extent.Text -match "'FullRun'" }, $true)
        $text = $(if ($branch) { $branch.Extent.Text } else { $script:Ast.Extent.Text })
        $start = $text.IndexOf("'OneClick' {")
        $start | Should -BeGreaterThan -1
        $oneClick = $text.Substring($start)
        $oneClick | Should -Match 'Get-MecmCurrentVersionByCMName'
        $oneClick | Should -Match 'Get-OneClickObservedSkipReason'
    }
}

Describe 'The site query connects to the provider that was asked for' {
    BeforeAll {
        function Get-CMApplication { param($Name, [switch]$DisableWildcardHandling, $ErrorAction) $script:Apps }
        function Set-Location { param($Path, $ErrorAction) $script:Located += [string]$Path }
        function Get-Command { param($Name, $ErrorAction) [pscustomobject]@{ Name = $Name } }
        function Get-PSDrive { param($Name, $PSProvider, $ErrorAction) if ($script:Drive) { $script:Drive } }
        function Remove-PSDrive { param($Name, [switch]$Force, $ErrorAction) $script:Removed += [string]$Name }
        function New-PSDrive { param($Name, $PSProvider, $Root, $ErrorAction) $script:Created += [string]$Root; [pscustomobject]@{ Name = $Name; Root = $Root } }
    }
    BeforeEach {
        $script:Located = @(); $script:Removed = @(); $script:Created = @()
        $script:Apps = @([pscustomobject]@{ LocalizedDisplayName = 'App'; Name = 'App'; SoftwareVersion = '1.0' })
    }

    It 'compares the drive root with the provider' {
        Test-MecmDriveRoot -Root 'cm01' -Provider 'cm01' | Should -BeTrue
        Test-MecmDriveRoot -Root 'CM01.lab.test' -Provider 'cm01' | Should -BeTrue
        Test-MecmDriveRoot -Root 'cm01' -Provider 'CM01.lab.test' | Should -BeTrue
        Test-MecmDriveRoot -Root 'cm02' -Provider 'cm01' | Should -BeFalse
        Test-MecmDriveRoot -Root 'cm01x' -Provider 'cm01' | Should -BeFalse
        Test-MecmDriveRoot -Root 'anything' -Provider '' | Should -BeTrue
    }

    It 'rebuilds an existing site drive that points at another provider' {
        $script:Drive = [pscustomobject]@{ Name = 'MCM'; Root = 'old-provider' }
        $null = Get-MecmCurrentVersionByCMName -SiteCode 'MCM' -ProviderMachineName 'new-provider' -CMName 'App' -ExactMatch
        $script:Removed | Should -Contain 'MCM'
        $script:Created | Should -Be @('new-provider')
    }

    It 'keeps an existing site drive that points at the provider' {
        $script:Drive = [pscustomobject]@{ Name = 'MCM'; Root = 'new-provider' }
        $null = Get-MecmCurrentVersionByCMName -SiteCode 'MCM' -ProviderMachineName 'new-provider' -CMName 'App' -ExactMatch
        $script:Removed.Count | Should -Be 0
        $script:Created.Count | Should -Be 0
    }
}
