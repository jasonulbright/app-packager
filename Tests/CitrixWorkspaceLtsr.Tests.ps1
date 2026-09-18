BeforeDiscovery {
    $script:LtsrCases = @(
        @{ Name = 'package-citrixworkspace-ltsr-x86';   Architecture = 'x86';   LocalSource = 'Optional'; Catalog = '$true';  Is64Bit = $false }
        @{ Name = 'package-citrixworkspace-ltsr-x64';   Architecture = 'x64';   LocalSource = 'Required'; Catalog = '$false'; Is64Bit = $true }
        @{ Name = 'package-citrixworkspace-ltsr-arm64'; Architecture = 'ARM64'; LocalSource = 'Required'; Catalog = '$false'; Is64Bit = $true }
    )
}

BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-citrixworkspace-ltsr-x86.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-CitrixWorkspaceLtsrFromCatalog' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))
    $script:CatalogBaseUrl = 'https://downloadplugins.citrix.com/ReceiverUpdates/Prod'

    function New-CatalogXml {
        param([string[]]$Installers)
        '<Catalog platform="win32"><Installers name="Receiver">' + ($Installers -join '') + '</Installers></Catalog>'
    }
    function New-InstallerXml {
        param([string]$Stream, [string]$Version, [string]$Hash = ('a' * 64))
        "<Installer><DownloadURL>/Receiver/Win/CitrixWorkspaceApp$Version.exe</DownloadURL><Hash>$Hash</Hash><Stream>$Stream</Stream><Version>$Version</Version></Installer>"
    }
    function Get-ScriptAssignment {
        param([string]$Path, [string]$Variable)
        $tk = $null; $er = $null
        $tree = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tk, [ref]$er)
        $node = $tree.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq "`$$Variable" }, $false)
        $node.Right.Extent.Text
    }
}

Describe 'Citrix Workspace LTSR catalog lookup' {
    It 'returns the newest LTSR entry and ignores the Current stream' {
        $xml = New-CatalogXml @(
            (New-InstallerXml -Stream 'Current' -Version '26.9.0.100')
            (New-InstallerXml -Stream 'LTSR' -Version '26.7.0.269' -Hash ('3D' * 32))
            (New-InstallerXml -Stream 'LTSR' -Version '25.3.10.12')
        )
        $release = Get-CitrixWorkspaceLtsrFromCatalog -Xml $xml
        $release.Version | Should -Be '26.7.0.269'
        $release.DownloadUrl | Should -Be 'https://downloadplugins.citrix.com/ReceiverUpdates/Prod/Receiver/Win/CitrixWorkspaceApp26.7.0.269.exe'
        $release.Sha256 | Should -Be ('3d' * 32)
    }

    It 'returns nothing when the catalog has no LTSR entry' {
        $xml = New-CatalogXml @((New-InstallerXml -Stream 'Current' -Version '26.9.0.100'))
        Get-CitrixWorkspaceLtsrFromCatalog -Xml $xml | Should -BeNullOrEmpty
    }

    It 'skips an LTSR entry without a valid SHA-256' {
        $xml = New-CatalogXml @(
            (New-InstallerXml -Stream 'LTSR' -Version '26.7.0.269' -Hash 'not-a-hash')
            (New-InstallerXml -Stream 'LTSR' -Version '25.3.10.12')
        )
        (Get-CitrixWorkspaceLtsrFromCatalog -Xml $xml).Version | Should -Be '25.3.10.12'
    }
}

Describe 'Citrix Workspace LTSR packager <Name>' -ForEach $LtsrCases {
    BeforeAll {
        $script:Path = Join-Path $PSScriptRoot "..\Packagers\$Name.ps1"
        $script:Text = Get-Content -LiteralPath $script:Path -Raw
    }

    It 'parses without errors' {
        $tk = $null; $er = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($script:Path, [ref]$tk, [ref]$er)
        $er | Should -BeNullOrEmpty
    }

    It 'names itself and its architecture' {
        Get-ScriptAssignment -Path $script:Path -Variable 'PackagerName' | Should -Be "`"$Name`""
        Get-ScriptAssignment -Path $script:Path -Variable 'Architecture' | Should -Be "`"$Architecture`""
    }

    It 'carries the expected LocalSource tag' {
        $script:Text | Should -Match "(?m)^LocalSource:\s*$LocalSource\s*$"
    }

    It 'downloads from the catalog only for the build the catalog serves' {
        Get-ScriptAssignment -Path $script:Path -Variable 'CatalogDownload' | Should -Be $Catalog
    }

    It 'checks the registry view the build writes its uninstall entry to' {
        $script:Text | Should -Match "Is64Bit\s+= \(\`$Architecture -ne 'x86'\)"
        ($Architecture -ne 'x86') | Should -Be $Is64Bit
    }
}

Describe 'Citrix Workspace LTSR installer selection' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\Packagers\AppPackagerCommon.psd1" -Force
        $script:Folder = Join-Path $TestDrive 'cwa'
        New-Item -ItemType Directory -Path $script:Folder -Force | Out-Null
        foreach ($n in 'CitrixWorkspaceApp.exe', 'CitrixWorkspaceFullInstaller (1).exe', 'CitrixWorkspaceApp_x64.exe', 'CitrixWorkspaceFullInstaller_ARM64.exe') {
            Set-Content -LiteralPath (Join-Path $script:Folder $n) -Value 'x'
        }
        function Get-Selection {
            param([string]$Name)
            $p = Join-Path $PSScriptRoot "..\Packagers\$Name.ps1"
            $filter = (Get-ScriptAssignment -Path $p -Variable 'InstallerFilter').Trim('"')
            $exclude = @(& ([scriptblock]::Create((Get-ScriptAssignment -Path $p -Variable 'InstallerExclude'))))
            @(Get-ChildItem -LiteralPath $script:Folder -Filter $filter -File |
                Where-Object { $n = $_.Name; -not @($exclude | Where-Object { $n -like $_ }) } |
                ForEach-Object Name | Sort-Object)
        }
    }

    It 'x86 takes the unsuffixed installers and skips the other builds' {
        Get-Selection -Name 'package-citrixworkspace-ltsr-x86' | Should -Be @('CitrixWorkspaceApp.exe', 'CitrixWorkspaceFullInstaller (1).exe')
    }

    It 'x64 takes only the x64 build' {
        Get-Selection -Name 'package-citrixworkspace-ltsr-x64' | Should -Be @('CitrixWorkspaceApp_x64.exe')
    }

    It 'ARM64 takes only the ARM64 build' {
        Get-Selection -Name 'package-citrixworkspace-ltsr-arm64' | Should -Be @('CitrixWorkspaceFullInstaller_ARM64.exe')
    }
}
