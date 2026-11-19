#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The PowerShell 7 packagers by release channel: LTS (MSI), Stable (MSIX) and Devolutions (third party).

.DESCRIPTION
    package-powershell7.ps1 follows the LTS line (7.6, Microsoft Learn "PowerShell Support Lifecycle") and never takes a release of
    another line. The names and the help of the three packagers say the channel.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:LtsPath = Join-Path $script:RepoRoot 'Packagers\package-powershell7.ps1'
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:LtsPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Select-PowerShell7LtsRelease' }, $false)
    if (-not $fn) { throw 'Select-PowerShell7LtsRelease not found' }
    . ([scriptblock]::Create($fn.Extent.Text))

    function New-FakeRelease {
        param([string]$Tag, [string]$Published, [string[]]$Assets, [switch]$Prerelease)
        [pscustomobject]@{
            tag_name = $Tag; published_at = $Published; draft = $false; prerelease = [bool]$Prerelease
            assets = @($Assets | ForEach-Object { [pscustomobject]@{ name = $_; browser_download_url = "https://example.test/$Tag/$_" } })
        }
    }
    $script:Msi = '^PowerShell-.+-win-x64\.msi$'
    function Msi([string]$v) { @("PowerShell-$v-win-x64.msi", "PowerShell-$v-win-x64.zip") }
}

Describe 'PowerShell 7 LTS release selection' {
    It 'takes the newest 7.6 release that has the MSI, not the newer Stable release without one' {
        $releases = @(
            (New-FakeRelease 'v7.7.0' '2026-11-10T10:00:00Z' @('PowerShell-7.7.0-win-x64.zip', 'PowerShell-7.7.0.msixbundle')),
            (New-FakeRelease 'v7.6.10' '2026-11-05T10:00:00Z' (Msi '7.6.10')),
            (New-FakeRelease 'v7.6.9' '2026-10-01T10:00:00Z' (Msi '7.6.9')))
        $r = Select-PowerShell7LtsRelease -Releases $releases -Line '7.6' -AssetPattern $script:Msi
        $r.Version | Should -Be '7.6.10'
        $r.DownloadUrl | Should -Be 'https://example.test/v7.6.10/PowerShell-7.6.10-win-x64.msi'
    }

    It 'never takes a newer line even when that line has an MSI' {
        $releases = @(
            (New-FakeRelease 'v7.8.0' '2027-03-01T10:00:00Z' (Msi '7.8.0')),
            (New-FakeRelease 'v7.6.10' '2026-11-05T10:00:00Z' (Msi '7.6.10')))
        (Select-PowerShell7LtsRelease -Releases $releases -Line '7.6' -AssetPattern $script:Msi).Version | Should -Be '7.6.10'
    }

    It 'never takes an older line' {
        $releases = @((New-FakeRelease 'v7.5.11' '2026-09-01T10:00:00Z' (Msi '7.5.11')), (New-FakeRelease 'v7.4.22' '2026-08-01T10:00:00Z' (Msi '7.4.22')))
        Select-PowerShell7LtsRelease -Releases $releases -Line '7.6' -AssetPattern $script:Msi | Should -BeNullOrEmpty
    }

    It 'does not take 7.60 or a 7.6 preview for the 7.6 line' {
        $releases = @(
            (New-FakeRelease 'v7.60.1' '2027-01-01T10:00:00Z' (Msi '7.60.1')),
            (New-FakeRelease 'v7.6.11-preview.1' '2026-11-12T10:00:00Z' (Msi '7.6.11-preview.1') -Prerelease),
            (New-FakeRelease 'v7.6.10' '2026-11-05T10:00:00Z' (Msi '7.6.10')))
        (Select-PowerShell7LtsRelease -Releases $releases -Line '7.6' -AssetPattern $script:Msi).Version | Should -Be '7.6.10'
    }

    It 'names the 7.6 release without an MSI that was passed over' {
        $releases = @(
            (New-FakeRelease 'v7.6.11' '2026-11-15T10:00:00Z' @('PowerShell-7.6.11-win-x64.zip')),
            (New-FakeRelease 'v7.6.10' '2026-11-05T10:00:00Z' (Msi '7.6.10')))
        $r = Select-PowerShell7LtsRelease -Releases $releases -Line '7.6' -AssetPattern $script:Msi
        $r.Version | Should -Be '7.6.10'
        $r.SkippedNewer | Should -Be @('7.6.11')
    }

    It 'accepts a release list that Windows PowerShell 5.1 nests in one object' {
        $inner = @((New-FakeRelease 'v7.6.10' '2026-11-05T10:00:00Z' (Msi '7.6.10')))
        (Select-PowerShell7LtsRelease -Releases @(, $inner) -Line '7.6' -AssetPattern $script:Msi).Version | Should -Be '7.6.10'
    }
}

Describe 'PowerShell 7 packagers named by release channel' {
    BeforeAll {
        $script:Lts = Get-Content -LiteralPath $script:LtsPath -Raw
        $script:Msix = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Packagers\package-powershell7msix.ps1') -Raw
        $script:Dev = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Packagers\package-devolutionspowershell.ps1') -Raw
    }

    It 'names the LTS packager LTS, keeps its file name and CMName, and pins the line' {
        $script:Lts | Should -Match '(?m)^App: PowerShell 7 LTS \(MSI\)\r?$'
        $script:Lts | Should -Match '(?m)^CMName: PowerShell 7\r?$'
        $script:Lts | Should -Match '\$LtsLine = "7\.6"'
        $script:Lts | Should -Match 'Long Term Servicing \(LTS\)'
        $script:Lts | Should -Match 'PowerShell Support Lifecycle'
        $script:Lts | Should -Match 'USE_MU=1'
    }

    It 'names the Stable MSIX packager Stable, per user, with the CMName that sites already use' {
        $script:Msix | Should -Match '(?m)^App: PowerShell 7 Stable \(MSIX, per user\)\r?$'
        $script:Msix | Should -Match '(?m)^CMName: PowerShell 7 \(MSIX\)\r?$'
        $script:Msix | Should -Match 'single-user installation'
        $script:Msix | Should -Match 'does not support PowerShell remoting'
        $script:Msix | Should -Match 'Stable release'
    }

    It 'names the Devolutions packager as third party' {
        $script:Dev | Should -Match '(?m)^App: PowerShell 7 \(Devolutions, third party\)\r?$'
        $script:Dev | Should -Match '(?m)^CMName: PowerShell 7 \(Devolutions\)\r?$'
        $script:Dev | Should -Match 'third party'
        $script:Dev | Should -Match 'Devolutions is not\s+Microsoft'
    }

    It 'tells each packager about the other two channels' {
        foreach ($text in $script:Lts, $script:Msix, $script:Dev) {
            $text | Should -Match 'package-powershell7msix\.ps1|PowerShell 7 Stable \(MSIX, per user\)'
        }
    }

    It 'gives all three one CPE so that the CVE lookup reads the version of each line' {
        foreach ($text in $script:Lts, $script:Msix, $script:Dev) {
            $text | Should -Match '(?m)^CPE: cpe:2\.3:a:microsoft:powershell:\*:\*:\*:\*:\*:\*:\*:\*\r?$'
        }
    }
}
