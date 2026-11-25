#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Check ConfigMgr finds the application that a packager creates, in four tiers, and picks the best fit when several match.

.DESCRIPTION
    Tier 1 the title equals CMName, tier 2 the CMName words with release details, tier 3 the title starts with the CMName words
    (the installer names it: "PuTTY release 0.85 (64-bit)"), tier 4 the Add/Remove Programs name of the packager ("Microsoft Azure PowerShell - *").
    The names of the applications that the packagers create do not change; only the search is wider. The ConfigMgr cmdlet is a stub that
    answers from a list, with the wildcard of the real -Name parameter.
#>

BeforeAll {
    $script:PackagersRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'Packagers'
    Import-Module (Join-Path $script:PackagersRoot 'AppPackagerDiscovery.psd1') -Force

    function global:Get-CMApplication {
        param([string]$Name, $ErrorAction)
        $pattern = New-Object System.Management.Automation.WildcardPattern($Name, [System.Management.Automation.WildcardOptions]::IgnoreCase)
        @($global:TierApps | Where-Object { $pattern.IsMatch($_.LocalizedDisplayName) })
    }
    function global:New-App([string]$Title, [string]$Version = '1.0') {
        [pscustomobject]@{ LocalizedDisplayName = $Title; Name = $Title; SoftwareVersion = $Version }
    }
}

AfterAll {
    Remove-Item function:\global:Get-CMApplication -ErrorAction SilentlyContinue
    Remove-Variable -Name TierApps -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Find-MecmApplicationByName' {
    It 'finds <Created> by CMName <CMName> in tier <Tier>' -ForEach @(
        @{ CMName = 'PuTTY'; Arp = 'PuTTY release * (64-bit)'; Created = 'PuTTY release 0.85 (64-bit)'; Tier = 'prefix' }
        @{ CMName = 'Azure PowerShell'; Arp = 'Microsoft Azure PowerShell - *'; Created = 'Microsoft Azure PowerShell - November 2026'; Tier = 'arp' }
        @{ CMName = 'Intune Debug Toolkit'; Arp = 'IntuneDebugTools*'; Created = 'IntuneDebugTools'; Tier = 'arp' }
        @{ CMName = 'Git'; Arp = ''; Created = 'Git 2.47.1 (x64)'; Tier = 'rule' }
        @{ CMName = 'Audacity'; Arp = ''; Created = 'Audacity'; Tier = 'name' }
    ) {
        $global:TierApps = @((New-App $Created '1.2'), (New-App 'GitHub Desktop' '3.4'))
        $hit = Find-MecmApplicationByName -CMName $CMName -ArpName $Arp
        $hit.Tier | Should -Be $Tier
        @($hit.Apps)[0].LocalizedDisplayName | Should -Be $Created
    }

    It 'finds nothing when no title carries the name' {
        $global:TierApps = @((New-App 'Microsoft Azure CLI'), (New-App 'Digit Tool'))
        (Find-MecmApplicationByName -CMName 'Git' -ArpName '').Tier | Should -Be ''
        (Find-MecmApplicationByName -CMName 'Azure PowerShell' -ArpName 'Microsoft Azure PowerShell - *').Tier | Should -Be ''
    }
}

Describe 'Select-MecmApplicationMatch' {
    It 'prefers the exact title' {
        $apps = @((New-App 'Audacity - 3.7.9' '3.7.9'), (New-App 'Audacity' '3.7.1'))
        (Select-MecmApplicationMatch -Apps $apps -CMName 'Audacity').LocalizedDisplayName | Should -Be 'Audacity'
    }
    It 'prefers PowerShell 7-x64 to PowerShell 7 (Devolutions) although the other is newer' {
        $apps = @((New-App 'PowerShell 7 (Devolutions)' '7.6.2'), (New-App 'PowerShell 7-x64' '7.6.1'))
        (Select-MecmApplicationMatch -Apps $apps -CMName 'PowerShell 7').LocalizedDisplayName | Should -Be 'PowerShell 7-x64'
    }
    It 'takes the newest version when the titles fit alike' {
        $apps = @((New-App 'NetBeans - 30' '30'), (New-App 'NetBeans - 31' '31'))
        (Select-MecmApplicationMatch -Apps $apps -CMName 'NetBeans').LocalizedDisplayName | Should -Be 'NetBeans - 31'
    }
    It 'prefers the title with the fewest extra words' {
        $apps = @((New-App 'Chrome Remote Desktop Host' '9.0'), (New-App 'Chrome Beta Channel' '2.0'))
        (Select-MecmApplicationMatch -Apps $apps -CMName 'Chrome').LocalizedDisplayName | Should -Be 'Chrome Beta Channel'
    }
}

Describe 'Get-MecmTitleExtraWordCount and Test-MecmTitleStartsWith' {
    It 'counts <Extra> extra word(s) after <CMName> in <Title>' -ForEach @(
        @{ CMName = 'PowerShell 7'; Title = 'PowerShell 7-x64'; Extra = 0 }
        @{ CMName = 'PowerShell 7'; Title = 'PowerShell 7 (Devolutions)'; Extra = 1 }
        @{ CMName = 'Git'; Title = 'Git 2.47.1 (x64)'; Extra = 0 }
        @{ CMName = 'Mozilla Firefox'; Title = 'Mozilla Firefox ESR (x64 en-US)'; Extra = 1 }
        @{ CMName = 'PuTTY'; Title = 'PuTTY release 0.85 (64-bit)'; Extra = 1 }
    ) {
        Get-MecmTitleExtraWordCount -CMName $CMName -Title $Title | Should -Be $Extra
    }
    It 'ends the word: <Title> for <CMName> is <Starts>' -ForEach @(
        @{ CMName = 'Git'; Title = 'GitHub Desktop'; Starts = $false }
        @{ CMName = 'Git'; Title = 'Git Extensions 4.3'; Starts = $true }
        @{ CMName = 'Git'; Title = 'Digit Tool'; Starts = $false }
        @{ CMName = 'Git'; Title = 'git'; Starts = $true }
    ) {
        Test-MecmTitleStartsWith -CMName $CMName -Title $Title | Should -Be $Starts
    }
}
