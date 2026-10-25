#Requires -Modules Pester

<#
.SYNOPSIS
    A rebuilt vendor installer is a different release.

.DESCRIPTION
    Git for Windows ships "v2.53.0.windows.2" as Git-2.53.0.2-64-bit.exe, and
    GIMP ships a rebuild as gimp-3.0.4-setup-2.exe. The resolver of each
    packager is lifted out by AST and run against a synthetic feed, so no
    vendor endpoint is contacted.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking

    function Get-FunctionFromPackager {
        param([string]$Name, [string]$Function)
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot "Packagers\package-$Name.ps1"), [ref]$null, [ref]$null)
        $node = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Function }, $true)
        if (-not $node) { throw "package-$Name.ps1 no longer defines $Function." }
        return $node.Extent.Text
    }
    function Get-AssignmentText {
        param([string]$Name, [string]$Variable)
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot "Packagers\package-$Name.ps1"), [ref]$null, [ref]$null)
        $node = $ast.Find({ param($n)
            $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.Left.VariablePath.UserPath -eq $Variable
        }, $true)
        return $node.Extent.Text
    }
    . ([scriptblock]::Create((Get-AssignmentText 'git' 'GitHubApiUrl')))
    . ([scriptblock]::Create((Get-FunctionFromPackager 'git' 'Get-LatestGitRelease')))
    . ([scriptblock]::Create((Get-AssignmentText 'gimp' 'CdnDirectoryUrl')))
    . ([scriptblock]::Create((Get-FunctionFromPackager 'gimp' 'Get-LatestGIMPVersion')))
    function Write-Log { param($Message, $Level, [switch]$Quiet) }
    function Get-GitHubApiCurlArgs { @() }
    $script:Feed = ''
    function curl.exe { $global:LASTEXITCODE = 0; $script:Feed }
}

Describe 'Git for Windows rebuilds' {
    It 'reports the first build of a release as the three-part version' {
        $script:Feed = '{"tag_name":"v2.53.0.windows.1","assets":[{"name":"Git-2.53.0-64-bit.exe","browser_download_url":"https://example.invalid/Git-2.53.0-64-bit.exe"}]}'
        (Get-LatestGitRelease).Version | Should -Be '2.53.0'
    }

    It 'reports a rebuild as a different version' {
        $script:Feed = '{"tag_name":"v2.53.0.windows.2","assets":[{"name":"Git-2.53.0.2-64-bit.exe","browser_download_url":"https://example.invalid/Git-2.53.0.2-64-bit.exe"}]}'
        $release = Get-LatestGitRelease
        $release.Version | Should -Be '2.53.0.2'
        $release.FileName | Should -Be 'Git-2.53.0.2-64-bit.exe'
    }

    It 'reads the build number from the installed registry value' {
        $text = [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot 'Packagers\package-git.ps1'))
        $text | Should -Match ([regex]::Escape('\.windows\.(?<build>\d+)$'))
        # The detection snippet's version normalization, run on registry values.
        $normalize = {
            param([string]$CurrentVersion)
            $v = $CurrentVersion.Trim()
            if ($v -match '^(?<base>\d+(\.\d+)*)\.windows\.(?<build>\d+)$') {
                $v = $Matches['base']
                if ([int]$Matches['build'] -gt 1) { $v = $v + '.' + $Matches['build'] }
            }
            $v
        }
        & $normalize '2.53.0.windows.1' | Should -Be '2.53.0'
        & $normalize '2.53.0.windows.2' | Should -Be '2.53.0.2'
        [version](& $normalize '2.53.0.windows.1') -lt [version](& $normalize '2.53.0.windows.2') | Should -BeTrue
    }
}

Describe 'GIMP rebuilds' {
    It 'keeps the installed version apart from the release identity' {
        $script:Feed = '<a href="0.0_LATEST-IS-3.0.4-2">marker</a>'
        $info = Get-LatestGIMPVersion
        $info.Version | Should -Be '3.0.4'
        $info.ReleaseVersion | Should -Be '3.0.4-2'
        $info.FileName | Should -Be 'gimp-3.0.4-setup-2.exe'
    }

    It 'gives the first build the plain version for both' {
        $script:Feed = '<a href="0.0_LATEST-IS-3.0.4-1">marker</a>'
        $info = Get-LatestGIMPVersion
        $info.Version | Should -Be '3.0.4'
        $info.ReleaseVersion | Should -Be '3.0.4'
        $info.FileName | Should -Be 'gimp-3.0.4-setup.exe'
    }

    It 'detects by the installed version and names the folder by the release' {
        $text = [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot 'Packagers\package-gimp.ps1'))
        $text | Should -Match ([regex]::Escape('ExpectedValue       = $installedVersion'))
        $text | Should -Match ([regex]::Escape('Write-Output $info.ReleaseVersion'))
    }
}

Describe 'NVIDIA PackageOnly' {

    It 'packages the staged build when the vendor lists a newer one' -ForEach @(
        @{ Name = 'nvidia-geforce'; Package = 'Invoke-PackageNvidiaGeForce'; Resolver = 'Resolve-NvidiaGeForceLatest' }
        @{ Name = 'nvidia-rtx-enterprise'; Package = 'Invoke-PackageNvidiaRTXEnterprise'; Resolver = 'Resolve-NvidiaRTXEnterpriseLatest' }
    ) {
        $text = Get-FunctionFromPackager $Name $Package
        $script:BaseDownloadRoot = Join-Path $TestDrive "nv-$Name"
        New-Item -ItemType Directory -Path (Join-Path $script:BaseDownloadRoot '550.1') -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:BaseDownloadRoot '550.1\stage-manifest.json') -Value '{}' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $script:BaseDownloadRoot 'staged-version.txt') -Value '550.1' -Encoding ASCII
        $global:NvidiaReadPath = $null
        $global:NvidiaResolverCalls = 0
        Set-Item -Path "function:global:$Resolver" -Value { $global:NvidiaResolverCalls++; [pscustomobject]@{ Version = '551.1' } }
        function global:Read-StageManifest { param($Path) $global:NvidiaReadPath = $Path; throw 'stop after the manifest path is chosen' }
        $BaseDownloadRoot = $script:BaseDownloadRoot
        try {
            . ([scriptblock]::Create($text))
            { & $Package } | Should -Throw '*stop after the manifest path is chosen*'
            $global:NvidiaReadPath | Should -Be (Join-Path $script:BaseDownloadRoot '550.1\stage-manifest.json')
            $global:NvidiaResolverCalls | Should -Be 0
        }
        finally {
            Remove-Item -Path "function:global:$Resolver", 'function:global:Read-StageManifest' -ErrorAction SilentlyContinue
        }
    }

    It 'asks for a stage when no marker exists' -ForEach @(
        @{ Name = 'nvidia-geforce'; Package = 'Invoke-PackageNvidiaGeForce' }
        @{ Name = 'nvidia-rtx-enterprise'; Package = 'Invoke-PackageNvidiaRTXEnterprise' }
    ) {
        $text = Get-FunctionFromPackager $Name $Package
        $BaseDownloadRoot = Join-Path $TestDrive "nv-empty-$Name"
        . ([scriptblock]::Create($text))
        { & $Package } | Should -Throw '*run Stage phase first*'
    }
}

Describe 'Spectra PDF PackageOnly' {
    It 'packages the staged build when the vendor lists a newer one' {
        $text = Get-FunctionFromPackager 'spectrapdf' 'Invoke-PackageSpectraPdf'
        $script:BaseDownloadRoot = Join-Path $TestDrive 'spectra'
        New-Item -ItemType Directory -Path (Join-Path $script:BaseDownloadRoot '1.0') -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:BaseDownloadRoot '1.0\stage-manifest.json') -Value '{}' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $script:BaseDownloadRoot 'staged-version.txt') -Value '1.0' -Encoding ASCII
        $global:SpectraReadPath = $null
        $global:SpectraResolverCalls = 0
        function global:Get-SpectraPdfLatestRelease { param([switch]$Quiet) $global:SpectraResolverCalls++; [pscustomobject]@{ Version = '2.0' } }
        function global:Read-StageManifest { param($Path) $global:SpectraReadPath = $Path; throw 'stop after the manifest path is chosen' }
        $BaseDownloadRoot = $script:BaseDownloadRoot
        try {
            . ([scriptblock]::Create($text))
            { Invoke-PackageSpectraPdf } | Should -Throw '*stop after the manifest path is chosen*'
            $global:SpectraReadPath | Should -Be (Join-Path $script:BaseDownloadRoot '1.0\stage-manifest.json')
            $global:SpectraResolverCalls | Should -Be 0
        }
        finally { Remove-Item -Path 'function:global:Get-SpectraPdfLatestRelease', 'function:global:Read-StageManifest' -ErrorAction SilentlyContinue }
    }
}

Describe 'Format templates' {
    It 'import the module from the folder they are copied into' {
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Packagers\Templates') -Filter '*.template') {
            $text = [System.IO.File]::ReadAllText($file.FullName)
            $text | Should -Match ([regex]::Escape('Import-Module "$PSScriptRoot\AppPackagerCommon.psd1"')) -Because $file.Name
            $text | Should -Not -Match ([regex]::Escape('$PSScriptRoot\..\AppPackagerCommon.psd1')) -Because $file.Name
        }
    }
}
