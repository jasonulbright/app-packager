#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Invoke-DiscoveryChoice reports a result that was not saved or not shown.
#>

BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-DiscoveryChoice' }, $false)
    if (-not $fn) { throw 'function Invoke-DiscoveryChoice not found' }
    . ([scriptblock]::Create($fn.Extent.Text))

    $script:Prefs = [pscustomobject]@{}
    function Enable-DiscoveredApplication { param($Prefs, $Script) @($Script).Count }
    function Add-DiscoveryDismissed { param($Prefs, $Script) @($Script).Count }
    function Add-LogLine { param([string]$Message) $script:LogLines += $Message }
    function Save-Preferences { param($Prefs) if ($script:SaveFails) { throw 'disk is full' } }
    function Invoke-RefreshGrid { if ($script:RefreshFails) { throw 'grid is busy' } }
}

Describe 'Invoke-DiscoveryChoice' {
    BeforeEach {
        $script:LogLines = @(); $script:SaveFails = $false; $script:RefreshFails = $false
        $script:txtStatus = [pscustomobject]@{ Text = '' }
        $txtStatus = $script:txtStatus
        $script:Choice = [pscustomobject]@{ Action = 'EnableAll'; Scripts = @('package-a.ps1', 'package-b.ps1') }
    }

    It 'reports the apps added when the choice is saved and shown' {
        Invoke-DiscoveryChoice -Choice $script:Choice
        $script:txtStatus.Text | Should -Be '2 app(s) added to the grid.'
    }

    It 'says the choice was not saved and where it holds' {
        $script:SaveFails = $true
        Invoke-DiscoveryChoice -Choice $script:Choice
        $script:txtStatus.Text | Should -Match 'not saved'
        $script:txtStatus.Text | Should -Match 'disk is full'
        ($script:LogLines -join ' ') | Should -Match 'not saved'
    }

    It 'says the grid was not refreshed' {
        $script:RefreshFails = $true
        Invoke-DiscoveryChoice -Choice $script:Choice
        $script:txtStatus.Text | Should -Match 'grid was not refreshed'
        $script:txtStatus.Text | Should -Match 'grid is busy'
        $script:txtStatus.Text | Should -Not -Match 'added to the grid\.$'
    }

    It 'says a dismissal was not saved' {
        $script:SaveFails = $true
        Invoke-DiscoveryChoice -Choice ([pscustomobject]@{ Action = 'Dismiss'; Scripts = @('package-a.ps1') })
        ($script:LogLines -join ' ') | Should -Match 'not saved'
        $script:txtStatus.Text | Should -Match 'not saved'
    }
}
