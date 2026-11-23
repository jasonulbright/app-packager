#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }

<#
.SYNOPSIS
    Pester 5.x guard: the window text of start-apppackager.ps1 never calls a device a client (owner term rule,
    lexicon LX-005, reader check 2026-11-23). The client is the Configuration Manager agent on a device.
#>

Describe 'Device and client wording' {
    It 'names devices where the text means the devices' -TestCases @(
        @{ Pattern = 'Clients install' }
        @{ Pattern = 'trusted on your clients' }
        @{ Pattern = 'no client gets it' }
    ) {
        $text = [System.IO.File]::ReadAllText((Join-Path (Split-Path -Parent $PSScriptRoot) 'start-apppackager.ps1'))
        $text | Should -Not -Match ([regex]::Escape($Pattern))
    }
}
