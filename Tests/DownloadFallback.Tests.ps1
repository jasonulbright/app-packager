#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    A vendor that serves one file under two addresses (Tableau: esdalt and tssoftware): the second address is used only when
    the host reports the first one as missing (404). A 403 is no answer and never switches the address.

.DESCRIPTION
    Select-AvailableDownloadUrl takes a probe script block, so no request goes to a host. Get-DownloadUrlState reads the
    answer of a real request; it runs against a local server.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:First = 'https://downloads.tableau.com/esdalt/tableau_prep/2026.3.0/TableauPrep-2026-3-0.exe'
    $script:Second = 'https://downloads.tableau.com/tssoftware/TableauPrep-2026-3-0.exe'
    function New-Probe { param([hashtable]$Table) $t = $Table; return { param($u) $global:ProbeAsked += $u; $t[$u] }.GetNewClosure() }
}

Describe 'Select-AvailableDownloadUrl' {
    BeforeEach { $global:ProbeAsked = @() }

    It 'takes the first address when it exists, and asks no further' {
        Select-AvailableDownloadUrl -Urls @($script:First, $script:Second) -Probe (New-Probe @{ $script:First = 'Exists'; $script:Second = 'Exists' }) | Should -Be $script:First
        $global:ProbeAsked | Should -Be @($script:First)
    }

    It 'takes the second address when the first one is reported as missing (404)' {
        Select-AvailableDownloadUrl -Urls @($script:First, $script:Second) -Probe (New-Probe @{ $script:First = 'Missing'; $script:Second = 'Exists' }) | Should -Be $script:Second
        $global:ProbeAsked | Should -Be @($script:First, $script:Second)
    }

    It 'stays with the first address when the host cannot tell (403) and never tries the second' {
        Select-AvailableDownloadUrl -Urls @($script:First, $script:Second) -Probe (New-Probe @{ $script:First = 'Unknown'; $script:Second = 'Exists' }) | Should -Be $script:First
        $global:ProbeAsked | Should -Be @($script:First)
    }

    It 'uses the second address even when it cannot tell about it, after a 404 on the first' {
        Select-AvailableDownloadUrl -Urls @($script:First, $script:Second) -Probe (New-Probe @{ $script:First = 'Missing'; $script:Second = 'Unknown' }) | Should -Be $script:Second
    }

    It 'stops with a clear message when every address is missing' {
        { Select-AvailableDownloadUrl -Urls @($script:First, $script:Second) -Probe (New-Probe @{ $script:First = 'Missing'; $script:Second = 'Missing' }) } |
            Should -Throw '*not on the host at any of its addresses (404)*esdalt*tssoftware*'
    }

    It 'ignores an empty entry in the list' {
        Select-AvailableDownloadUrl -Urls @('', $script:Second) -Probe (New-Probe @{ $script:Second = 'Exists' }) | Should -Be $script:Second
    }
}

Describe 'Get-DownloadUrlState' {
    BeforeAll {
        $script:Port = Get-Random -Minimum 20000 -Maximum 40000
        $script:Listener = New-Object System.Net.HttpListener
        $script:Listener.Prefixes.Add("http://localhost:$($script:Port)/")
        $script:Listener.Start()
        $script:Ps = [powershell]::Create()
        [void]$script:Ps.AddScript({
                param($listener)
                while ($listener.IsListening) {
                    try { $ctx = $listener.GetContext() } catch { break }
                    $path = $ctx.Request.Url.AbsolutePath
                    $ctx.Response.StatusCode = switch ($path) { '/ok' { 206 } '/missing' { 404 } '/gone' { 410 } '/forbidden' { 403 } default { 500 } }
                    $ctx.Response.Close()
                }
            }).AddArgument($script:Listener)
        $script:Handle = $script:Ps.BeginInvoke()
        $script:Base = "http://localhost:$($script:Port)"
        Mock -ModuleName AppPackagerCommon Write-Log { }
    }
    AfterAll {
        try { $script:Listener.Stop(); $script:Listener.Close() } catch { $null = $_ }
        try { $script:Ps.Dispose() } catch { $null = $_ }
    }

    It 'reads 200 and 206 as Exists, 404 and 410 as Missing, and every other answer as Unknown' {
        Get-DownloadUrlState -Url "$($script:Base)/ok" | Should -Be 'Exists'
        Get-DownloadUrlState -Url "$($script:Base)/missing" | Should -Be 'Missing'
        Get-DownloadUrlState -Url "$($script:Base)/gone" | Should -Be 'Missing'
        Get-DownloadUrlState -Url "$($script:Base)/forbidden" | Should -Be 'Unknown'
        Get-DownloadUrlState -Url "$($script:Base)/other" | Should -Be 'Unknown'
    }

    It 'reads a host that does not answer as Unknown, never as Missing' {
        Get-DownloadUrlState -Url "http://localhost:$(Get-Random -Minimum 41000 -Maximum 45000)/x" | Should -Be 'Unknown'
    }
}

Describe 'The Tableau packagers use the fallback' {
    It 'select the address before the download in every Tableau packager' -ForEach @('package-tableaudesktop.ps1', 'package-tableaureader.ps1', 'package-tableauprep.ps1') {
        $text = Get-Content -LiteralPath (Join-Path $PSScriptRoot "..\Packagers\$_") -Raw
        $text | Should -Match 'Select-AvailableDownloadUrl -Urls @\(\$release\.DownloadUrl, \$release\.FallbackUrl\)'
        $text | Should -Match 'Invoke-DownloadWithRetry -Url \$downloadUrl'
        $text | Should -Match '\$DownloadUrlTemplate = "https://downloads\.tableau\.com/esdalt/'
        $text | Should -Match '\$FallbackUrlTemplate = "https://downloads\.tableau\.com/tssoftware/'
    }
}
