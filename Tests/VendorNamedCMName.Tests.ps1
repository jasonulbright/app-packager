#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The header CMName of a packager finds the application that its Package phase creates, for packagers that
    take the name from the installer.

.DESCRIPTION
    The names below are the MSI ProductName values read from the installers on 2026-11-24 (PuTTY 0.85 and
    IntuneDebugTools 2.3). Get-MecmCurrentVersionByCMName finds an application by
    Test-MecmApplicationTitle, which needs every word of the CMName in order.
#>

BeforeAll {
    $script:PackagersRoot = Join-Path (Split-Path -Parent $PSScriptRoot) 'Packagers'
    Import-Module (Join-Path $script:PackagersRoot 'AppPackagerDiscovery.psd1') -Force
    function Get-HeaderCMName {
        param([string]$Packager)
        foreach ($l in (Get-Content -LiteralPath (Join-Path $script:PackagersRoot "package-$Packager.ps1") -TotalCount 40)) {
            if ($l -match '^\s*CMName\s*:\s*(.+?)\s*$') { return $Matches[1] }
        }
    }
}

Describe 'CMName of packagers that take the application name from the installer' {
    It 'finds <Created> with the CMName of package-<Packager>' -ForEach @(
        @{ Packager = 'putty'; Created = 'PuTTY release 0.85 (64-bit)' }
        @{ Packager = 'intunedebugtoolkit'; Created = 'IntuneDebugTools' }
    ) {
        $cmName = Get-HeaderCMName -Packager $Packager
        Test-MecmApplicationTitle -CMName $cmName -Title $Created | Should -BeTrue -Because "CMName '$cmName'"
    }
}
