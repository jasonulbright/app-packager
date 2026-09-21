#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x tests for AppPackagerCommon shared module.

.DESCRIPTION
    Tests pure-logic and local-filesystem functions. Does NOT require ConfigMgr,
    network shares, real MSI files, or administrator elevation.

.EXAMPLE
    Invoke-Pester .\AppPackagerCommon.Tests.ps1
#>

BeforeDiscovery {
    Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
}

BeforeAll {
    Import-Module "$PSScriptRoot\AppPackagerCommon.psd1" -Force
}

Describe 'Module exports' {
    It 'exports the MSIX wrapper helper advertised by templates' {
        Get-Command -Module AppPackagerCommon -Name New-MsixWrapperContent -ErrorAction SilentlyContinue |
            Should -Not -BeNullOrEmpty
    }
}

Describe 'Connect-CMSite' {
    # The wrapper keeps resolution and the fail-fast guard in
    # AppPackagerCommon; the drive and session mechanics execute inside
    # SuiteCommon. Mocks target the module that runs each piece.
    BeforeEach {
        $script:__savedSuiteProvider = $env:SUITE_CM_PROVIDER
        Remove-Item Env:SUITE_CM_PROVIDER -ErrorAction SilentlyContinue
    }
    AfterEach {
        if ($script:__savedSuiteProvider) { $env:SUITE_CM_PROVIDER = $script:__savedSuiteProvider }
    }

    It 'creates the CMSite PSDrive when the provider machine is configured' {
        Mock Get-Module   -ModuleName SuiteCommon { [pscustomobject]@{ Name = 'ConfigurationManager' } } -ParameterFilter { $Name -eq 'ConfigurationManager' }
        Mock Get-PSDrive  -ModuleName SuiteCommon { $null } -ParameterFilter { $Name -eq 'MCM' -and $PSProvider -eq 'CMSite' }
        Mock New-PSDrive  -ModuleName SuiteCommon { [pscustomobject]@{ Name = $Name; Root = $Root } } -ParameterFilter { $Name -eq 'MCM' -and $PSProvider -eq 'CMSite' -and $Root -eq 'provider.example' }
        Mock Set-Location -ModuleName SuiteCommon { }
        Mock Write-Log    -ModuleName SuiteCommon { }
        Mock Get-PSDrive  -ModuleName AppPackagerCommon { $null }

        Connect-CMSite -SiteCode 'MCM' -ProviderMachineName 'provider.example' | Should -BeTrue

        Should -Invoke New-PSDrive -ModuleName SuiteCommon -Times 1 -Exactly -ParameterFilter { $Name -eq 'MCM' -and $PSProvider -eq 'CMSite' -and $Root -eq 'provider.example' }
        Should -Invoke Set-Location -ModuleName SuiteCommon -Times 1 -Exactly -ParameterFilter { $Path -eq 'MCM:' }
    }

    It 'fails clearly when no drive exists and no provider machine can be resolved' {
        Mock Resolve-CMProviderMachineName -ModuleName AppPackagerCommon { $null }
        Mock Get-PSDrive  -ModuleName AppPackagerCommon { $null }
        Mock New-PSDrive  -ModuleName SuiteCommon { throw 'should not be called' }
        Mock Write-Log    -ModuleName AppPackagerCommon { }

        Connect-CMSite -SiteCode 'MCM' | Should -BeFalse

        Should -Invoke New-PSDrive -ModuleName SuiteCommon -Times 0 -Exactly
        Should -Invoke Write-Log -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $Level -eq 'ERROR' -and $Message -match 'no provider machine name is configured' }
    }

    It 'recreates the drive from its old root when entering an existing drive fails' {
        $global:__apEnterAttempts = 0
        Mock Get-Module   -ModuleName SuiteCommon { [pscustomobject]@{ Name = 'ConfigurationManager' } } -ParameterFilter { $Name -eq 'ConfigurationManager' }
        Mock Get-PSDrive  -ModuleName SuiteCommon { [pscustomobject]@{ Name = 'MCM'; Root = 'provider.example' } } -ParameterFilter { $Name -eq 'MCM' -and $PSProvider -eq 'CMSite' }
        Mock Set-Location -ModuleName SuiteCommon {
            if ($Path -eq 'MCM:') {
                $global:__apEnterAttempts += 1
                if ($global:__apEnterAttempts -eq 1) { throw 'drive has no provider connection' }
            }
        }
        Mock Remove-PSDrive -ModuleName SuiteCommon { }
        Mock New-PSDrive  -ModuleName SuiteCommon { [pscustomobject]@{ Name = $Name; Root = $Root } } -ParameterFilter { $Name -eq 'MCM' -and $PSProvider -eq 'CMSite' }
        Mock Write-Log    -ModuleName SuiteCommon { }
        Mock Resolve-CMProviderMachineName -ModuleName AppPackagerCommon { $null }
        Mock Get-PSDrive  -ModuleName AppPackagerCommon { [pscustomobject]@{ Name = 'MCM'; Root = 'provider.example' } }

        Connect-CMSite -SiteCode 'MCM' | Should -BeTrue

        Should -Invoke Remove-PSDrive -ModuleName SuiteCommon -Times 1 -Exactly -ParameterFilter { $Name -eq 'MCM' }
        Should -Invoke New-PSDrive -ModuleName SuiteCommon -Times 1 -Exactly -ParameterFilter { $Root -eq 'provider.example' }
        Remove-Variable -Name __apEnterAttempts -Scope Global -ErrorAction SilentlyContinue
    }
}

# ============================================================================
# Write-Log / Initialize-Logging
# ============================================================================

Describe 'Write-Log' {
    It 'writes formatted message to log file' {
        $logFile = Join-Path $TestDrive 'test.log'
        Initialize-Logging -LogPath $logFile

        Write-Log 'Hello world' -Quiet

        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[INFO \] Hello world'
    }

    It 'tags WARN messages correctly' {
        $logFile = Join-Path $TestDrive 'warn.log'
        Initialize-Logging -LogPath $logFile

        Write-Log 'Something odd' -Level WARN -Quiet

        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[WARN \] Something odd'
    }

    It 'tags ERROR messages correctly' {
        $logFile = Join-Path $TestDrive 'error.log'
        Initialize-Logging -LogPath $logFile

        Write-Log 'Failure' -Level ERROR -Quiet

        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[ERROR\] Failure'
    }

    It 'accepts empty string message' {
        $logFile = Join-Path $TestDrive 'empty.log'
        Initialize-Logging -LogPath $logFile

        { Write-Log '' -Quiet } | Should -Not -Throw

        $lines = Get-Content -LiteralPath $logFile
        # Header line + empty-message line
        $lines.Count | Should -BeGreaterOrEqual 2
    }
}

Describe 'Initialize-Logging' {
    It 'creates log file with header line' {
        $logFile = Join-Path $TestDrive 'init.log'
        Initialize-Logging -LogPath $logFile

        Test-Path -LiteralPath $logFile | Should -BeTrue
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[INFO \] === Log initialized ==='
    }

    It 'creates parent directories if missing' {
        $logFile = Join-Path $TestDrive 'sub\dir\deep.log'
        Initialize-Logging -LogPath $logFile

        Test-Path -LiteralPath $logFile | Should -BeTrue
    }
}

# ============================================================================
# New-MsiWrapperContent
# ============================================================================

Describe 'New-MsiWrapperContent' {
    BeforeAll {
        $result = New-MsiWrapperContent -MsiFileName 'acme-widget.msi'
    }

    It 'returns a hashtable with Install and Uninstall keys' {
        $result | Should -BeOfType [hashtable]
        $result.Keys | Should -Contain 'Install'
        $result.Keys | Should -Contain 'Uninstall'
    }

    It 'install script references the MSI filename' {
        $result.Install | Should -Match 'acme-widget\.msi'
    }

    It 'install script uses msiexec /i with /qn /norestart' {
        $result.Install | Should -Match 'msiexec\.exe'
        $result.Install | Should -Match '/i'
        $result.Install | Should -Match '/qn'
        $result.Install | Should -Match '/norestart'
    }

    It 'uninstall script uses msiexec /x with /qn /norestart' {
        $result.Uninstall | Should -Match 'msiexec\.exe'
        $result.Uninstall | Should -Match '/x'
        $result.Uninstall | Should -Match '/qn'
        $result.Uninstall | Should -Match '/norestart'
    }

    It 'both scripts exit with the captured msiexec code' {
        $result.Install   | Should -Match '\$exit = \$proc\.ExitCode'
        $result.Install   | Should -Match 'exit \$exit'
        $result.Uninstall | Should -Match 'exit \$proc\.ExitCode'
    }

    It 'uses an array ArgumentList (not a single string)' {
        $result.Install | Should -Match '\$args = @\('
        $result.Install | Should -Match '-ArgumentList \$args'
    }
}

# ============================================================================
# New-ExeWrapperContent
# ============================================================================

Describe 'New-ExeWrapperContent' {
    Context 'with uninstall args' {
        BeforeAll {
            $result = New-ExeWrapperContent `
                -InstallerFileName 'setup.exe' `
                -InstallArgs "'/S', '/norestart'" `
                -UninstallCommand 'C:\Program Files\Acme\uninstall.exe' `
                -UninstallArgs "'/S'"
        }

        It 'returns a hashtable with Install and Uninstall keys' {
            $result | Should -BeOfType [hashtable]
            $result.Keys | Should -Contain 'Install'
            $result.Keys | Should -Contain 'Uninstall'
        }

        It 'install script references the installer filename' {
            $result.Install | Should -Match 'setup\.exe'
        }

        It 'install script includes install args' {
            $result.Install | Should -Match '/S'
        }

        It 'uninstall script references the uninstall command' {
            $result.Uninstall | Should -Match 'uninstall\.exe'
        }

        It 'uninstall script includes uninstall args' {
            $result.Uninstall | Should -Match '/S'
        }

        It 'both scripts end with exit $proc.ExitCode' {
            $result.Install   | Should -Match 'exit \$proc\.ExitCode'
            $result.Uninstall | Should -Match 'exit \$proc\.ExitCode'
        }
    }

    Context 'without uninstall args' {
        BeforeAll {
            $result = New-ExeWrapperContent `
                -InstallerFileName 'setup.exe' `
                -InstallArgs "'/S'" `
                -UninstallCommand 'C:\Program Files\Acme\uninstall.exe'
        }

        It 'uninstall script omits -ArgumentList when args empty' {
            $result.Uninstall | Should -Not -Match '-ArgumentList'
        }
    }

    Context 'environment variables in the uninstall path' {
        It 'expands %VAR% segments at run time instead of passing them literally' {
            $result = New-ExeWrapperContent -InstallerFileName 'setup.exe' -InstallArgs "'/S'" `
                -UninstallCommand '%LOCALAPPDATA%\Acme\Uninstall.exe' -UninstallArgs "'/S'"
            $result.Uninstall | Should -Match 'ExpandEnvironmentVariables\(''%LOCALAPPDATA%\\Acme\\Uninstall\.exe''\)'
            $result.Uninstall | Should -Match '-FilePath \$uninstallPath '
            $expanded = [Environment]::ExpandEnvironmentVariables('%LOCALAPPDATA%\Acme\Uninstall.exe')
            $expanded | Should -Be (Join-Path $env:LOCALAPPDATA 'Acme\Uninstall.exe')
        }

        It 'keeps an apostrophe in the path inside the literal' {
            $result = New-ExeWrapperContent -InstallerFileName 'setup.exe' -InstallArgs "'/S'" `
                -UninstallCommand "C:\Program Files\O'Neil\uninstall.exe"
            $result.Uninstall | Should -Match "O''Neil"
            $null = [scriptblock]::Create($result.Uninstall)
        }
    }

    Context 'processes the installer launches on success' {
        It 'waits for the installer alone and closes the launched processes' {
            $result = New-ExeWrapperContent -InstallerFileName 'setup.exe' -InstallArgs "'/S'" `
                -UninstallCommand 'C:\Program Files\Acme\uninstall.exe' -PostInstallKillProcesses @('Acme', 'AcmeTray')
            $result.Install | Should -Match '\$proc\.WaitForExit\(\)'
            $result.Install | Should -Not -Match '-Wait'
            $result.Install | Should -Match "@\('Acme', 'AcmeTray'\)"
            $result.Install | Should -Match 'Stop-Process -Force'
            $result.Install | Should -Match 'exit \$exit'
            $null = [scriptblock]::Create($result.Install)
        }

        It 'keeps the descendant-aware wait when no process is named' {
            $result = New-ExeWrapperContent -InstallerFileName 'setup.exe' -InstallArgs "'/S'" -UninstallCommand 'x.exe'
            $result.Install | Should -Match '-Wait -PassThru'
            $result.Install | Should -Not -Match 'Stop-Process'
        }
    }
}

# ============================================================================
# Write-ContentWrappers
# ============================================================================

Describe 'Write-ContentWrappers' {
    BeforeAll {
        $outDir = Join-Path $TestDrive 'wrappers'
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null

        Write-ContentWrappers `
            -OutputPath $outDir `
            -InstallPs1Content 'echo install' `
            -UninstallPs1Content 'echo uninstall'
    }

    It 'creates install.bat' {
        Test-Path (Join-Path $outDir 'install.bat') | Should -BeTrue
    }

    It 'creates install.ps1' {
        Test-Path (Join-Path $outDir 'install.ps1') | Should -BeTrue
    }

    It 'creates uninstall.bat' {
        Test-Path (Join-Path $outDir 'uninstall.bat') | Should -BeTrue
    }

    It 'creates uninstall.ps1' {
        Test-Path (Join-Path $outDir 'uninstall.ps1') | Should -BeTrue
    }

    It 'install.bat contains @echo off' {
        $bat = Get-Content (Join-Path $outDir 'install.bat') -Raw
        $bat | Should -Match '@echo off'
    }

    It 'install.bat calls PowerShell.exe with install.ps1' {
        $bat = Get-Content (Join-Path $outDir 'install.bat') -Raw
        $bat | Should -Match 'PowerShell\.exe.*install\.ps1'
    }

    It 'install.bat propagates ERRORLEVEL by default' {
        $bat = Get-Content (Join-Path $outDir 'install.bat') -Raw
        $bat | Should -Match 'exit /b %ERRORLEVEL%'
    }

    It 'install.ps1 contains the provided content' {
        $ps1 = Get-Content (Join-Path $outDir 'install.ps1') -Raw
        $ps1 | Should -Match 'echo install'
    }

    It 'uninstall.ps1 contains the provided content' {
        $ps1 = Get-Content (Join-Path $outDir 'uninstall.ps1') -Raw
        $ps1 | Should -Match 'echo uninstall'
    }

    It 'overwrites existing files on second call' {
        # Overwrite install.ps1 with custom content, then verify wrapper
        # generation returns the staged content to its deterministic form.
        Set-Content (Join-Path $outDir 'install.ps1') -Value 'custom' -Encoding ASCII

        Write-ContentWrappers `
            -OutputPath $outDir `
            -InstallPs1Content 'NEW content' `
            -UninstallPs1Content 'NEW uninstall'

        $ps1 = Get-Content (Join-Path $outDir 'install.ps1') -Raw
        $ps1 | Should -Match 'NEW content'
        $ps1 | Should -Not -Match 'custom'
    }

    Context 'custom bat exit codes' {
        BeforeAll {
            $customDir = Join-Path $TestDrive 'custom-exit'
            New-Item -ItemType Directory -Path $customDir -Force | Out-Null

            Write-ContentWrappers `
                -OutputPath $customDir `
                -InstallPs1Content 'echo install' `
                -UninstallPs1Content 'echo uninstall' `
                -InstallBatExitCode '3010' `
                -UninstallBatExitCode '0'
        }

        It 'install.bat uses custom exit code 3010' {
            $bat = Get-Content (Join-Path $customDir 'install.bat') -Raw
            $bat | Should -Match 'exit /b 3010'
        }

        It 'uninstall.bat uses custom exit code 0' {
            $bat = Get-Content (Join-Path $customDir 'uninstall.bat') -Raw
            $bat | Should -Match 'exit /b 0'
        }
    }
}

# ============================================================================
# Write-StageManifest / Read-StageManifest
# ============================================================================

Describe 'Write-StageManifest' {
    It 'writes valid JSON with SchemaVersion and StagedAt' {
        $stageDir = Join-Path $TestDrive 'manifest-basic'
        New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $stageDir 'setup.msi') -Value 'payload' -Encoding ASCII
        $path = Join-Path $stageDir 'manifest.json'

        Write-StageManifest -Path $path -ManifestData @{
            AppName         = 'Test App - 1.0'
            Publisher       = 'Test Vendor'
            SoftwareVersion = '1.0'
        }

        Test-Path -LiteralPath $path | Should -BeTrue
        $json = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $json.SchemaVersion | Should -Be 4
        $json.StagedAt | Should -Not -BeNullOrEmpty
        $json.AppName | Should -Be 'Test App - 1.0'
        $json.Publisher | Should -Be 'Test Vendor'
        $json.FileHashes | Should -Not -BeNullOrEmpty
    }

    It 'populates FileHashes for payloads and wrappers and excludes the manifest file itself' {
        $stageDir = Join-Path $TestDrive 'manifest-hashes'
        New-Item -ItemType Directory -Path $stageDir -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $stageDir 'setup.exe') -Value 'payload bytes' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $stageDir 'install.ps1') -Value 'install wrapper' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $stageDir 'uninstall.ps1') -Value 'uninstall wrapper' -Encoding ASCII

        $path = Join-Path $stageDir 'stage-manifest.json'
        Write-StageManifest -Path $path -ManifestData @{
            AppName         = 'Hash App - 1.0'
            Publisher       = 'Test Vendor'
            SoftwareVersion = '1.0'
        }

        $manifest = Read-StageManifest -Path $path
        $relativePaths = @($manifest.FileHashes | ForEach-Object { [string]$_.RelativePath })

        $relativePaths | Should -Contain 'setup.exe'
        $relativePaths | Should -Contain 'install.ps1'
        $relativePaths | Should -Contain 'uninstall.ps1'
        $relativePaths | Should -Not -Contain 'stage-manifest.json'
        foreach ($entry in @($manifest.FileHashes)) {
            ([string]$entry.Sha256) | Should -Match '^[A-F0-9]{64}$'
            [int64]$entry.Size | Should -BeGreaterThan 0
        }
    }
}

Describe 'Read-StageManifest' {
    It 'round-trips manifest data correctly' {
        $path = Join-Path $TestDrive 'roundtrip.json'

        $data = @{
            AppName         = 'RoundTrip App - 2.5'
            Publisher       = 'Acme Corp'
            SoftwareVersion = '2.5.0'
            InstallerFile   = 'setup.msi'
            Detection       = @{
                Type                = 'RegistryKeyValue'
                RegistryKeyRelative = 'SOFTWARE\Test\Key'
                ValueName           = 'DisplayVersion'
                ExpectedValue       = '2.5.0.0'
                Operator            = 'IsEquals'
                Is64Bit             = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.AppName         | Should -Be 'RoundTrip App - 2.5'
        $manifest.Publisher       | Should -Be 'Acme Corp'
        $manifest.SoftwareVersion | Should -Be '2.5.0'
        $manifest.InstallerFile   | Should -Be 'setup.msi'
        $manifest.Detection.Type  | Should -Be 'RegistryKeyValue'
        $manifest.Detection.RegistryKeyRelative | Should -Be 'SOFTWARE\Test\Key'
        $manifest.Detection.ExpectedValue | Should -Be '2.5.0.0'
        $manifest.Detection.Operator | Should -Be 'IsEquals'
        $manifest.Detection.Is64Bit | Should -BeTrue
        $manifest.PSObject.Properties.Name | Should -Contain 'FileHashes'
    }

    It 'throws when file does not exist' {
        { Read-StageManifest -Path (Join-Path $TestDrive 'nonexistent.json') } |
            Should -Throw '*not found*'
    }

    It 'throws when JSON is missing SchemaVersion' {
        $path = Join-Path $TestDrive 'bad-manifest.json'
        '{"AppName": "test"}' | Set-Content -LiteralPath $path -Encoding UTF8

        { Read-StageManifest -Path $path } |
            Should -Throw '*missing SchemaVersion*'
    }

    It 'soft-lands pre-1.0.7 manifests without FileHashes' {
        $path = Join-Path $TestDrive 'old-manifest.json'
        @{
            SchemaVersion   = 2
            AppName         = 'Old App - 1.0'
            Publisher       = 'Legacy'
            SoftwareVersion = '1.0'
        } | ConvertTo-Json | Set-Content -LiteralPath $path -Encoding UTF8

        { $script:oldManifest = Read-StageManifest -Path $path } | Should -Not -Throw
        $script:oldManifest.SchemaVersion | Should -Be 2
        $script:oldManifest.PSObject.Properties.Name | Should -Not -Contain 'FileHashes'
    }
}

Describe 'Compare-StageFileHashes' {
    BeforeEach {
        $script:hashRoot = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $script:hashRoot -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:hashRoot 'payload.bin') -Value 'payload' -Encoding ASCII
        Set-Content -LiteralPath (Join-Path $script:hashRoot 'install.ps1') -Value 'install' -Encoding ASCII
        $nested = Join-Path $script:hashRoot 'nested'
        New-Item -ItemType Directory -Path $nested -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $nested 'data.txt') -Value 'nested-data' -Encoding ASCII
        $script:expectedHashes = Get-StageFileHashes -Root $script:hashRoot
    }

    It 'passes when the tree matches the recorded hashes' {
        $result = Compare-StageFileHashes -Root $script:hashRoot -Expected $script:expectedHashes

        $result.Pass | Should -BeTrue
        $result.ExpectedCount | Should -Be 3
        $result.ActualCount | Should -Be 3
    }

    It "fails when a file's bytes change" {
        Set-Content -LiteralPath (Join-Path $script:hashRoot 'payload.bin') -Value 'tampered' -Encoding ASCII

        $result = Compare-StageFileHashes -Root $script:hashRoot -Expected $script:expectedHashes

        $result.Pass | Should -BeFalse
        $result.Mismatches | Should -HaveCount 1
        $result.Mismatches[0].RelativePath | Should -Be 'payload.bin'
    }

    It 'fails when an expected file is missing' {
        Remove-Item -LiteralPath (Join-Path $script:hashRoot 'install.ps1') -Force

        $result = Compare-StageFileHashes -Root $script:hashRoot -Expected $script:expectedHashes

        $result.Pass | Should -BeFalse
        $result.Missing | Should -HaveCount 1
        $result.Missing[0].RelativePath | Should -Be 'install.ps1'
    }

    It 'fails when an unexpected extra file is present unless extras are allowed' {
        Set-Content -LiteralPath (Join-Path $script:hashRoot 'extra.txt') -Value 'extra' -Encoding ASCII

        $strict = Compare-StageFileHashes -Root $script:hashRoot -Expected $script:expectedHashes
        $strict.Pass | Should -BeFalse
        $strict.Extra | Should -HaveCount 1
        $strict.Extra[0].RelativePath | Should -Be 'extra.txt'

        $allowed = Compare-StageFileHashes -Root $script:hashRoot -Expected $script:expectedHashes -AllowExtra
        $allowed.Pass | Should -BeTrue
    }

    It 'skips verification for missing expected hashes to support pre-1.0.7 manifests' {
        $result = Compare-StageFileHashes -Root $script:hashRoot -Expected $null

        $result.Pass | Should -BeTrue
        $result.Skipped | Should -BeTrue
        $result.Reason | Should -Match 'does not contain FileHashes'
    }
}

# ============================================================================
# New-OdtConfigXml
# ============================================================================

Describe 'New-OdtConfigXml' {
    Context 'basic single-product download XML' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '64' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail') `
                -SourcePath 'C:\temp\ap\M365Apps-x64\16.0.19127.20532'
        }

        It 'starts with <Configuration>' {
            $xml | Should -Match '^<Configuration>'
        }

        It 'ends with closing Configuration tag' {
            $xml.TrimEnd() | Should -BeLike '*</Configuration>'
        }

        It 'includes OfficeClientEdition 64' {
            $xml | Should -Match 'OfficeClientEdition="64"'
        }

        It 'includes the version' {
            $xml | Should -Match 'Version="16\.0\.19127\.20532"'
        }

        It 'includes Channel MonthlyEnterprise (default)' {
            $xml | Should -Match 'Channel="MonthlyEnterprise"'
        }

        It 'includes the SourcePath' {
            $xml | Should -Match 'SourcePath="C:\\temp\\ap\\M365Apps-x64\\16\.0\.19127\.20532"'
        }

        It 'includes the product ID' {
            $xml | Should -Match 'Product ID="O365ProPlusRetail"'
        }

        It 'excludes Groove, Lync, OneDrive, Teams, Bing' {
            $xml | Should -Match 'ExcludeApp ID="Groove"'
            $xml | Should -Match 'ExcludeApp ID="Lync"'
            $xml | Should -Match 'ExcludeApp ID="OneDrive"'
            $xml | Should -Match 'ExcludeApp ID="Teams"'
            $xml | Should -Match 'ExcludeApp ID="Bing"'
        }

        It 'includes SharedComputerLicensing' {
            $xml | Should -Match 'Name="SharedComputerLicensing" Value="1"'
        }

        It 'includes FORCEAPPSHUTDOWN' {
            $xml | Should -Match 'Name="FORCEAPPSHUTDOWN" Value="TRUE"'
        }

        It 'includes MigrateArch' {
            $xml | Should -Match 'MigrateArch="TRUE"'
        }

        It 'includes RemoveMSI' {
            $xml | Should -Match '<RemoveMSI />'
        }

        It 'includes Display Level None with AcceptEULA' {
            $xml | Should -Match 'Display Level="None" AcceptEULA="TRUE"'
        }

        It 'includes Logging element' {
            $xml | Should -Match 'Logging Level="Standard"'
        }
    }

    Context 'install XML without SourcePath' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '64' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail')
        }

        It 'does not include SourcePath attribute' {
            $xml | Should -Not -Match 'SourcePath='
        }
    }

    Context 'multi-product XML' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '64' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail', 'VisioProRetail')
        }

        It 'includes both product IDs' {
            $xml | Should -Match 'Product ID="O365ProPlusRetail"'
            $xml | Should -Match 'Product ID="VisioProRetail"'
        }

        It 'each product has its own ExcludeApp entries' {
            # Two sets of ExcludeApp blocks (one per product)
            $grooveMatches = [regex]::Matches($xml, 'ExcludeApp ID="Groove"')
            $grooveMatches.Count | Should -Be 2
        }
    }

    Context 'with CompanyName' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '32' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail') `
                -CompanyName 'Contoso Ltd'
        }

        It 'includes AppSettings block' {
            $xml | Should -Match '<AppSettings>'
        }

        It 'includes Company setup with the provided name' {
            $xml | Should -Match 'Name="Company" Value="Contoso Ltd"'
        }
    }

    Context 'with XML-sensitive preference values' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '64' `
                -ProductIds @('O365ProPlusRetail') `
                -CompanyName 'A&B <Lab>' `
                -SourcePath 'C:\Temp\Office & Apps'
        }

        It 'produces parseable XML' {
            { [xml]$xml } | Should -Not -Throw
        }

        It 'escapes company and source path attributes' {
            $xml | Should -Match 'Value="A&amp;B &lt;Lab&gt;"'
            $xml | Should -Match 'SourcePath="C:\\Temp\\Office &amp; Apps"'
        }
    }

    Context 'without CompanyName' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '64' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail')
        }

        It 'omits AppSettings block entirely' {
            $xml | Should -Not -Match '<AppSettings>'
            $xml | Should -Not -Match 'Name="Company"'
        }
    }

    Context 'x86 edition' {
        BeforeAll {
            $xml = New-OdtConfigXml `
                -OfficeClientEdition '32' `
                -Version '16.0.19127.20532' `
                -ProductIds @('O365ProPlusRetail')
        }

        It 'includes OfficeClientEdition 32' {
            $xml | Should -Match 'OfficeClientEdition="32"'
        }
    }
}

# ============================================================================
# Initialize-Folder
# ============================================================================

Describe 'Initialize-Folder' {
    It 'creates a new directory' {
        $dir = Join-Path $TestDrive 'new-folder'
        Initialize-Folder -Path $dir

        Test-Path -LiteralPath $dir | Should -BeTrue
        (Get-Item $dir).PSIsContainer | Should -BeTrue
    }

    It 'does not error when directory already exists' {
        $dir = Join-Path $TestDrive 'existing-folder'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null

        { Initialize-Folder -Path $dir } | Should -Not -Throw
    }

    It 'creates nested directories' {
        $dir = Join-Path $TestDrive 'a\b\c'
        Initialize-Folder -Path $dir

        Test-Path -LiteralPath $dir | Should -BeTrue
    }
}

# ============================================================================
# Get-PackagerPreferences
# ============================================================================

Describe 'Get-PackagerPreferences' {
    It 'reads the actual packager-preferences.json file' {
        $prefsPath = Join-Path $PSScriptRoot 'packager-preferences.json'
        if (-not (Test-Path -LiteralPath $prefsPath)) {
            Set-ItResult -Skipped -Because 'packager-preferences.json not present'
            return
        }

        $prefs = Get-PackagerPreferences
        $prefs | Should -Not -BeNullOrEmpty
        $prefs.PSObject.Properties.Name | Should -Contain 'CompanyName'
    }
}

# ============================================================================
# Write-StageManifest / Read-StageManifest - per-user manifest overrides
# ============================================================================

Describe 'Stage manifest with per-user deployment overrides' {
    It 'round-trips InstallationBehaviorType and LogonRequirementType' {
        $path = Join-Path $TestDrive 'zoom-manifest.json'

        $data = @{
            AppName                  = 'Zoom Workplace - 6.6.0 (x64)'
            Publisher                = 'Zoom Video Communications'
            SoftwareVersion          = '6.6.0'
            InstallerFile            = 'ZoomInstaller.exe'
            InstallationBehaviorType = 'InstallForUser'
            LogonRequirementType     = 'OnlyWhenUserLoggedOn'
            Detection                = @{
                Type         = 'File'
                FilePath     = '%APPDATA%\Zoom\bin'
                FileName     = 'Zoom.exe'
                PropertyType = 'Existence'
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.InstallationBehaviorType | Should -Be 'InstallForUser'
        $manifest.LogonRequirementType     | Should -Be 'OnlyWhenUserLoggedOn'
        $manifest.Detection.Type           | Should -Be 'File'
        $manifest.Detection.FilePath       | Should -Be '%APPDATA%\Zoom\bin'
        $manifest.Detection.FileName       | Should -Be 'Zoom.exe'
        $manifest.Detection.PropertyType   | Should -Be 'Existence'
    }
}

# ============================================================================
# Write-StageManifest - RegistryKeyValue with fixed ARP key
# ============================================================================

Describe 'Stage manifest with fixed ARP key detection' {
    It 'round-trips RegistryKeyValue detection with named key' {
        $path = Join-Path $TestDrive 'vlc-manifest.json'

        $data = @{
            AppName         = 'VLC Media Player - 3.0.23 (x64)'
            Publisher       = 'VideoLAN'
            SoftwareVersion = '3.0.23'
            InstallerFile   = 'vlc-3.0.23-win64.msi'
            Detection       = @{
                Type                = 'RegistryKeyValue'
                RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VLC media player'
                ValueName           = 'DisplayVersion'
                ExpectedValue       = '3.0.23'
                Operator            = 'IsEquals'
                Is64Bit             = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.Detection.Type | Should -Be 'RegistryKeyValue'
        $manifest.Detection.RegistryKeyRelative | Should -Match 'VLC media player'
        $manifest.Detection.ExpectedValue | Should -Be '3.0.23'
        $manifest.Detection.Operator | Should -Be 'IsEquals'
    }
}

# ============================================================================
# Write-StageManifest - File Existence with version-specific path (R pattern)
# ============================================================================

Describe 'Stage manifest with version-specific file detection path' {
    It 'round-trips File Existence detection with R-style versioned path' {
        $path = Join-Path $TestDrive 'r-manifest.json'

        $data = @{
            AppName         = 'R for Windows - 4.5.2 (x64)'
            Publisher       = 'The R Foundation'
            SoftwareVersion = '4.5.2'
            InstallerFile   = 'R-4.5.2-win.exe'
            Detection       = @{
                Type         = 'File'
                FilePath     = 'C:\Program Files\R\R-4.5.2\bin'
                FileName     = 'R.exe'
                PropertyType = 'Existence'
                Is64Bit      = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.AppName         | Should -Be 'R for Windows - 4.5.2 (x64)'
        $manifest.Detection.Type  | Should -Be 'File'
        $manifest.Detection.FilePath     | Should -Be 'C:\Program Files\R\R-4.5.2\bin'
        $manifest.Detection.FileName     | Should -Be 'R.exe'
        $manifest.Detection.PropertyType | Should -Be 'Existence'
    }
}

# ============================================================================
# Write-StageManifest - RegistryKeyValue with '+' in version (RStudio pattern)
# ============================================================================

Describe 'Stage manifest with plus sign in version string' {
    It 'round-trips RegistryKeyValue detection preserving + in ExpectedValue' {
        $path = Join-Path $TestDrive 'rstudio-manifest.json'

        $data = @{
            AppName         = 'RStudio Desktop - 2026.01.1+403 (x64)'
            Publisher       = 'Posit Software, PBC'
            SoftwareVersion = '2026.01.1+403'
            InstallerFile   = 'RStudio-2026.01.1-403.exe'
            Detection       = @{
                Type                = 'RegistryKeyValue'
                RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\RStudio'
                ValueName           = 'DisplayVersion'
                ExpectedValue       = '2026.01.1+403'
                Operator            = 'IsEquals'
                Is64Bit             = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.SoftwareVersion              | Should -Be '2026.01.1+403'
        $manifest.Detection.ExpectedValue      | Should -Be '2026.01.1+403'
        $manifest.Detection.RegistryKeyRelative | Should -Match 'RStudio'
        $manifest.Detection.Operator           | Should -Be 'IsEquals'
    }
}

# ============================================================================
# Write-StageManifest - File Existence with dash-build version (Positron)
# ============================================================================

Describe 'Stage manifest with dash-build version string' {
    It 'round-trips File Existence detection with Positron-style version' {
        $path = Join-Path $TestDrive 'positron-manifest.json'

        $data = @{
            AppName         = 'Positron - 2026.02.1-5 (x64)'
            Publisher       = 'Posit Software, PBC'
            SoftwareVersion = '2026.02.1-5'
            InstallerFile   = 'Positron-2026.02.1-5-Setup-x64.exe'
            Detection       = @{
                Type         = 'File'
                FilePath     = 'C:\Program Files\Positron'
                FileName     = 'Positron.exe'
                PropertyType = 'Existence'
                Is64Bit      = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.SoftwareVersion        | Should -Be '2026.02.1-5'
        $manifest.Detection.Type         | Should -Be 'File'
        $manifest.Detection.FilePath     | Should -Be 'C:\Program Files\Positron'
        $manifest.Detection.FileName     | Should -Be 'Positron.exe'
    }
}

# ============================================================================
# Write-StageManifest - File Existence with version-specific folder (Python)
# ============================================================================

Describe 'Stage manifest with Python-style version-specific install path' {
    It 'round-trips File Existence detection with Python314 folder path' {
        $path = Join-Path $TestDrive 'python-manifest.json'

        $data = @{
            AppName         = 'Python - 3.14.3 (x64)'
            Publisher       = 'Python Software Foundation'
            SoftwareVersion = '3.14.3'
            InstallerFile   = 'python-3.14.3-amd64.exe'
            Detection       = @{
                Type         = 'File'
                FilePath     = 'C:\Program Files\Python314'
                FileName     = 'python.exe'
                PropertyType = 'Existence'
                Is64Bit      = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.SoftwareVersion        | Should -Be '3.14.3'
        $manifest.Detection.FilePath     | Should -Be 'C:\Program Files\Python314'
        $manifest.Detection.FileName     | Should -Be 'python.exe'
    }
}

# ============================================================================
# Write-StageManifest - File Existence with ProgramData path (Anaconda)
# ============================================================================

Describe 'Stage manifest with ProgramData detection path' {
    It 'round-trips File Existence detection with Anaconda ProgramData path' {
        $path = Join-Path $TestDrive 'anaconda-manifest.json'

        $data = @{
            AppName         = 'Anaconda Distribution - 2025.12-2 (x64)'
            Publisher       = 'Anaconda, Inc.'
            SoftwareVersion = '2025.12-2'
            InstallerFile   = 'Anaconda3-2025.12-2-Windows-x86_64.exe'
            Detection       = @{
                Type         = 'File'
                FilePath     = 'C:\ProgramData\anaconda3'
                FileName     = 'python.exe'
                PropertyType = 'Existence'
                Is64Bit      = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.SoftwareVersion        | Should -Be '2025.12-2'
        $manifest.Detection.FilePath     | Should -Be 'C:\ProgramData\anaconda3'
        $manifest.Detection.FileName     | Should -Be 'python.exe'
    }
}

# ============================================================================
# Write-StageManifest - Temurin JRE 8 ARP detection with + in version
# ============================================================================

Describe 'Stage manifest with Temurin JRE 8 ARP detection' {
    It 'round-trips RegistryKeyValue detection with +build version' {
        $path = Join-Path $TestDrive 'temurin-jre8-manifest.json'

        $data = @{
            AppName         = 'Eclipse Temurin JRE 8 - 8.0.482+8 (x64)'
            Publisher       = 'Eclipse Adoptium'
            SoftwareVersion = '8.0.482+8'
            InstallerFile   = 'OpenJDK8U-jre_x64_windows_hotspot_8u482b08.msi'
            Detection       = @{
                Type                = 'RegistryKeyValue'
                RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{A8C9D8D3-7E2A-4B1F-8C4E-12345678ABCD}'
                ValueName           = 'DisplayVersion'
                ExpectedValue       = '8.0.482.8'
                Is64Bit             = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.AppName         | Should -Be 'Eclipse Temurin JRE 8 - 8.0.482+8 (x64)'
        $manifest.Publisher       | Should -Be 'Eclipse Adoptium'
        $manifest.SoftwareVersion | Should -Be '8.0.482+8'
        $manifest.Detection.Type  | Should -Be 'RegistryKeyValue'
        $manifest.Detection.ExpectedValue | Should -Be '8.0.482.8'
        $manifest.Detection.Is64Bit | Should -BeTrue
    }
}

# ============================================================================
# Write-StageManifest - Temurin JDK 21 ARP detection
# ============================================================================

Describe 'Stage manifest with Temurin JDK 21 ARP detection' {
    It 'round-trips RegistryKeyValue detection with standard version' {
        $path = Join-Path $TestDrive 'temurin-jdk21-manifest.json'

        $data = @{
            AppName         = 'Eclipse Temurin JDK 21 - 21.0.10+7 (x64)'
            Publisher       = 'Eclipse Adoptium'
            SoftwareVersion = '21.0.10+7'
            InstallerFile   = 'OpenJDK21U-jdk_x64_windows_hotspot_21.0.10_7.msi'
            Detection       = @{
                Type                = 'RegistryKeyValue'
                RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{B2D4E6F8-1234-5678-9ABC-DEF012345678}'
                ValueName           = 'DisplayVersion'
                ExpectedValue       = '21.0.10.7'
                Is64Bit             = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.AppName         | Should -Be 'Eclipse Temurin JDK 21 - 21.0.10+7 (x64)'
        $manifest.SoftwareVersion | Should -Be '21.0.10+7'
        $manifest.Detection.RegistryKeyRelative | Should -Match 'B2D4E6F8'
        $manifest.Detection.ExpectedValue | Should -Be '21.0.10.7'
        $manifest.Detection.Is64Bit | Should -BeTrue
    }
}

# ============================================================================
# Write-StageManifest - Corretto JDK 21 ARP detection (4-part normalized)
# ============================================================================

Describe 'Stage manifest with Corretto JDK 21 ARP detection' {
    It 'round-trips RegistryKeyValue detection with 4-part normalized version' {
        $path = Join-Path $TestDrive 'corretto-jdk21-manifest.json'

        $data = @{
            AppName         = 'Amazon Corretto JDK 21 - 21.0.10.7 (x64)'
            Publisher       = 'Amazon'
            SoftwareVersion = '21.0.10.7'
            InstallerFile   = 'amazon-corretto-21.0.10.7.1-windows-x64.msi'
            Detection       = @{
                Type                = 'RegistryKeyValue'
                RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{C3E5F7A9-ABCD-1234-5678-90ABCDEF1234}'
                ValueName           = 'DisplayVersion'
                ExpectedValue       = '21.0.10.7'
                Is64Bit             = $true
            }
        }

        Write-StageManifest -Path $path -ManifestData $data
        $manifest = Read-StageManifest -Path $path

        $manifest.AppName         | Should -Be 'Amazon Corretto JDK 21 - 21.0.10.7 (x64)'
        $manifest.Publisher       | Should -Be 'Amazon'
        $manifest.SoftwareVersion | Should -Be '21.0.10.7'
        $manifest.Detection.Type  | Should -Be 'RegistryKeyValue'
        $manifest.Detection.ExpectedValue | Should -Be '21.0.10.7'
        $manifest.Detection.Is64Bit | Should -BeTrue
    }
}

# ============================================================================
# Schema v2 - PSADT / deployment tool integration fields
# ============================================================================

Describe 'Schema v2: MSI manifest with install/uninstall/process fields' {
    BeforeAll {
        $script:v2Path = Join-Path $TestDrive 'v2-msi-manifest.json'

        Write-StageManifest -Path $v2Path -ManifestData @{
            AppName         = '7-Zip - 26.00 (x64)'
            Publisher       = 'Igor Pavlov'
            SoftwareVersion = '26.00'
            InstallerFile   = '7z2600-x64.msi'
            InstallerType   = 'MSI'
            InstallArgs     = '/qn /norestart'
            UninstallArgs   = '/qn /norestart'
            ProductCode     = '{23170F69-40C1-2702-2600-000001000000}'
            RunningProcess  = @('7zFM', '7zG')
            Detection       = @{
                Type                = 'RegistryKeyValue'
                RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{23170F69-40C1-2702-2600-000001000000}'
                ValueName           = 'DisplayVersion'
                DisplayVersion      = '26.00.00.0'
                Is64Bit             = $true
            }
        }
        $script:v2Manifest = Read-StageManifest -Path $v2Path
    }

    It 'emits SchemaVersion 3' {
        $v2Manifest.SchemaVersion | Should -Be 4
    }

    It 'includes InstallerType' {
        $v2Manifest.InstallerType | Should -Be 'MSI'
    }

    It 'includes InstallArgs' {
        $v2Manifest.InstallArgs | Should -Be '/qn /norestart'
    }

    It 'includes UninstallArgs' {
        $v2Manifest.UninstallArgs | Should -Be '/qn /norestart'
    }

    It 'includes ProductCode' {
        $v2Manifest.ProductCode | Should -Be '{23170F69-40C1-2702-2600-000001000000}'
    }

    It 'includes RunningProcess as array' {
        $v2Manifest.RunningProcess | Should -HaveCount 2
        $v2Manifest.RunningProcess | Should -Contain '7zFM'
        $v2Manifest.RunningProcess | Should -Contain '7zG'
    }

    It 'still includes all v1 fields' {
        $v2Manifest.AppName         | Should -Not -BeNullOrEmpty
        $v2Manifest.Publisher       | Should -Not -BeNullOrEmpty
        $v2Manifest.SoftwareVersion | Should -Not -BeNullOrEmpty
        $v2Manifest.InstallerFile   | Should -Not -BeNullOrEmpty
        $v2Manifest.Detection       | Should -Not -BeNullOrEmpty
        $v2Manifest.StagedAt        | Should -Not -BeNullOrEmpty
    }
}

Describe 'Schema v2: EXE manifest with UninstallCommand and RunningProcess' {
    BeforeAll {
        $script:v2ExePath = Join-Path $TestDrive 'v2-exe-manifest.json'

        Write-StageManifest -Path $v2ExePath -ManifestData @{
            AppName          = 'PyCharm Community - 2025.2.6'
            Publisher        = 'JetBrains'
            SoftwareVersion  = '2025.2.6'
            InstallerFile    = 'pycharm-community-2025.2.6.exe'
            InstallerType    = 'EXE'
            InstallArgs      = '/S'
            UninstallCommand = 'C:\Program Files\JetBrains\PyCharm Community Edition 2025.2.6\bin\Uninstall.exe'
            UninstallArgs    = '/S'
            RunningProcess   = @('pycharm64')
            Detection        = @{
                Type         = 'File'
                FilePath     = 'C:\Program Files\JetBrains\PyCharm Community Edition 2025.2.6\bin'
                FileName     = 'pycharm64.exe'
                PropertyType = 'Existence'
                Is64Bit      = $true
            }
        }
        $script:v2ExeManifest = Read-StageManifest -Path $v2ExePath
    }

    It 'includes InstallerType EXE' {
        $v2ExeManifest.InstallerType | Should -Be 'EXE'
    }

    It 'includes UninstallCommand for EXE products' {
        $v2ExeManifest.UninstallCommand | Should -Match 'Uninstall\.exe'
    }

    It 'includes InstallArgs' {
        $v2ExeManifest.InstallArgs | Should -Be '/S'
    }

    It 'includes RunningProcess' {
        $v2ExeManifest.RunningProcess | Should -Contain 'pycharm64'
    }
}

Describe 'Schema v2: backward compatibility with v1 manifests (no v2 fields)' {
    It 'v1-style manifest without v2 fields round-trips cleanly' {
        $path = Join-Path $TestDrive 'v1-compat.json'

        Write-StageManifest -Path $path -ManifestData @{
            AppName         = 'Legacy App - 1.0'
            Publisher       = 'Legacy Corp'
            SoftwareVersion = '1.0'
            InstallerFile   = 'setup.msi'
            Detection       = @{
                Type                = 'RegistryKeyValue'
                RegistryKeyRelative = 'SOFTWARE\Legacy\App'
                ValueName           = 'DisplayVersion'
                ExpectedValue       = '1.0'
            }
        }

        $manifest = Read-StageManifest -Path $path

        # optional installer/process fields should be absent (not null-filled or defaulted)
        $manifest.PSObject.Properties.Name | Should -Not -Contain 'InstallerType'
        $manifest.PSObject.Properties.Name | Should -Not -Contain 'InstallArgs'
        $manifest.PSObject.Properties.Name | Should -Not -Contain 'RunningProcess'

        # v1 fields still work
        $manifest.AppName | Should -Be 'Legacy App - 1.0'
        $manifest.SchemaVersion | Should -Be 4
        $manifest.PSObject.Properties.Name | Should -Contain 'FileHashes'
    }
}

# ============================================================================
# Get-NetworkContentPath
# ============================================================================

Describe 'Get-NetworkContentPath' {
    It 'builds and creates the nested layout by default' {
        $root = Join-Path $TestDrive 'share-nested'
        New-Item -ItemType Directory -Path $root -Force | Out-Null

        $path = Get-NetworkContentPath -FileServerPath $root -VendorFolder 'Igor Pavlov' -AppFolder '7-Zip' -Version '25.01'
        $path | Should -Be (Join-Path $root 'Applications\Igor Pavlov\7-Zip\25.01')
        Test-Path -LiteralPath $path | Should -BeTrue
    }

    It 'builds and creates the flat layout' {
        $root = Join-Path $TestDrive 'share-flat'
        New-Item -ItemType Directory -Path $root -Force | Out-Null

        $path = Get-NetworkContentPath -FileServerPath $root -VendorFolder 'Igor Pavlov' -AppFolder '7-Zip' -Version '25.01' -Layout Flat
        $path | Should -Be (Join-Path $root 'Applications\Igor Pavlov-7-Zip-25.01')
        Test-Path -LiteralPath $path | Should -BeTrue
        (Get-ChildItem -LiteralPath (Join-Path $root 'Applications') -Directory).Count | Should -Be 1
    }
}

# ============================================================================
# Test-PsadtLayout
# ============================================================================

Describe 'Test-PsadtLayout' {
    It 'detects a v4 layout with the exe launcher' {
        $root = Join-Path $TestDrive 'psadt-v4'
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        Set-Content -Path (Join-Path $root 'Invoke-AppDeployToolkit.ps1') -Value '# v4'
        Set-Content -Path (Join-Path $root 'Invoke-AppDeployToolkit.exe') -Value 'stub'

        $layout = Test-PsadtLayout -Path $root
        $layout.Generation | Should -Be 'v4'
        $layout.EntryPoint | Should -Be 'Invoke-AppDeployToolkit.exe'
        $layout.InstallCommandLine | Should -Be 'Invoke-AppDeployToolkit.exe -DeploymentType Install'
        $layout.UninstallCommandLine | Should -Be 'Invoke-AppDeployToolkit.exe -DeploymentType Uninstall'
    }

    It 'falls back to powershell.exe for a v4 layout without the exe' {
        $root = Join-Path $TestDrive 'psadt-v4-noexe'
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        Set-Content -Path (Join-Path $root 'Invoke-AppDeployToolkit.ps1') -Value '# v4'

        $layout = Test-PsadtLayout -Path $root
        $layout.Generation | Should -Be 'v4'
        $layout.InstallCommandLine | Should -Match '^powershell\.exe .*Invoke-AppDeployToolkit\.ps1" -DeploymentType Install$'
    }

    It 'detects a v3 layout and appends DeployMode when forced' {
        $root = Join-Path $TestDrive 'psadt-v3'
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        Set-Content -Path (Join-Path $root 'Deploy-Application.exe') -Value 'stub'
        Set-Content -Path (Join-Path $root 'Deploy-Application.ps1') -Value '# v3'

        $layout = Test-PsadtLayout -Path $root -DeployMode Silent
        $layout.Generation | Should -Be 'v3'
        $layout.EntryPoint | Should -Be 'Deploy-Application.exe'
        $layout.InstallCommandLine | Should -Be 'Deploy-Application.exe -DeploymentType "Install" -DeployMode Silent'
    }

    It 'throws when no toolkit entry point exists' {
        $root = Join-Path $TestDrive 'psadt-empty'
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        { Test-PsadtLayout -Path $root } | Should -Throw '*No PSADT entry point*'
    }
}

# ============================================================================
# Get-NextPatchVersion
# ============================================================================

Describe 'Get-NextPatchVersion' {
    It 'increments the last numeric component' {
        Get-NextPatchVersion -Version '8.0.20' | Should -Be '8.0.21'
        Get-NextPatchVersion -Version '10.0.1' | Should -Be '10.0.2'
        Get-NextPatchVersion -Version '9' | Should -Be '10'
    }

    It 'returns null for non-numeric last components' {
        Get-NextPatchVersion -Version '10.0.0-rc.1' | Should -BeNullOrEmpty
        Get-NextPatchVersion -Version '8.0.20-preview' | Should -BeNullOrEmpty
    }
}

# ============================================================================
# New-MECMApplicationFromManifest - existing application validation
# ============================================================================

Describe 'New-MECMApplicationFromManifest existing application validation' {
    InModuleScope AppPackagerCommon {
        BeforeAll {
            function Get-CMApplication { }
            function Get-CMDeploymentType { }
            function New-CMApplication { param($Name, $Publisher, $SoftwareVersion, $Description, $AutoInstall, $LocalizedApplicationName, $ErrorAction) }
            function Add-CMScriptDeploymentType { param($ApplicationName, $DeploymentTypeName, $ContentLocation, $InstallCommand, $UninstallCommand, $InstallationBehaviorType, $LogonRequirementType, $EstimatedRuntimeMins, $MaximumRuntimeMins, [switch]$ContentFallback, $SlowNetworkDeploymentMode, $UserInteractionMode, $RebootBehavior, $ScriptLanguage, $ScriptText, $AddDetectionClause, $DetectionClauseConnector, $GroupDetectionClauses, [switch]$RequireUserInteraction, $ErrorAction) }
            function Remove-CMDeploymentType { param($ApplicationName, $DeploymentTypeName, [switch]$Force, $ErrorAction) }
            function Set-CMDeploymentType { param($ApplicationName, $DeploymentTypeName, $NewDeploymentTypeName, $ErrorAction) }
            function Set-CMApplication { param($Name, $SoftwareVersion, $Description, $Publisher, $AutoInstall, $ErrorAction) }
            # Detection clause construction runs unmocked; without the console
            # module loaded the cmdlet does not exist, so a stub returning a
            # Setting.LogicalName-shaped object stands in.
            function New-CMDetectionClauseFile {
                param($Path, $FileName, [switch]$Existence, [switch]$Is64Bit, $PropertyType, $ExpectedValue, $ExpressionOperator, [switch]$Value, $ErrorAction)
                [pscustomobject]@{ Setting = [pscustomobject]@{ LogicalName = 'File_' + [guid]::NewGuid().ToString() } }
            }
        }

        AfterAll {
            Remove-Item -Path function:\Get-CMApplication -ErrorAction SilentlyContinue
            Remove-Item -Path function:\Get-CMDeploymentType -ErrorAction SilentlyContinue
            Remove-Item -Path function:\New-CMApplication -ErrorAction SilentlyContinue
            Remove-Item -Path function:\Add-CMScriptDeploymentType -ErrorAction SilentlyContinue
            Remove-Item -Path function:\Remove-CMDeploymentType -ErrorAction SilentlyContinue
            Remove-Item -Path function:\Set-CMDeploymentType -ErrorAction SilentlyContinue
            Remove-Item -Path function:\Set-CMApplication -ErrorAction SilentlyContinue
            Remove-Item -Path function:\New-CMDetectionClauseFile -ErrorAction SilentlyContinue
        }

        BeforeEach {
            $script:testManifest = [pscustomobject]@{
                AppName         = 'Test App - 1.0'
                Publisher       = 'Contoso'
                SoftwareVersion = '1.0'
                Detection       = [pscustomobject]@{
                    Type       = 'Script'
                    ScriptText = 'Write-Output "Installed"'
                }
            }

            Mock Connect-CMSite { $true }
            Mock Get-CMApplication { [pscustomobject]@{ CI_ID = 1234; SoftwareVersion = '1.0' } }
            Mock Get-CMDeploymentType { [pscustomobject]@{ LocalizedDisplayName = 'Test App - 1.0' } }
        }

        It 'sets task-sequence eligibility for new applications: <Label>' -TestCases @(
            @{ Label = 'default system'; Behavior = $null; Logon = $null; Interaction = $false; Expected = $true }
            @{ Label = 'explicit system'; Behavior = 'InstallForSystem'; Logon = 'WhetherOrNotUserLoggedOn'; Interaction = $false; Expected = $true }
            @{ Label = 'user'; Behavior = 'InstallForUser'; Logon = 'OnlyWhenUserLoggedOn'; Interaction = $false; Expected = $false }
            @{ Label = 'user without logon override'; Behavior = 'InstallForUser'; Logon = $null; Interaction = $false; Expected = $false }
            @{ Label = 'conditional user'; Behavior = 'InstallForSystemIfResourceIsDeviceOtherwiseInstallForUser'; Logon = $null; Interaction = $false; Expected = $false }
            @{ Label = 'requires logged-on user'; Behavior = 'InstallForSystem'; Logon = 'OnlyWhenUserLoggedOn'; Interaction = $false; Expected = $false }
            @{ Label = 'requires interaction'; Behavior = 'InstallForSystem'; Logon = $null; Interaction = $true; Expected = $false }
        ) {
            param($Label, $Behavior, $Logon, $Interaction, $Expected)
            $script:testManifest | Add-Member -NotePropertyMembers @{
                InstallationBehaviorType = $Behavior
                LogonRequirementType = $Logon
                RequireUserInteraction = $Interaction
            }
            Mock Get-CMApplication { $null }
            Mock New-CMApplication { [pscustomobject]@{ CI_ID = 4321 } }
            Mock Add-CMScriptDeploymentType { }
            Mock Remove-CMApplicationRevisionHistoryByCIId { }
            New-MECMApplicationFromManifest -Manifest $script:testManifest -SiteCode MCM -NetworkContentPath '\\server\share\Test' | Should -Be 4321
            Should -Invoke New-CMApplication -Times 1 -Exactly -ParameterFilter { $AutoInstall -eq $Expected }
            if ($Behavior) {
                Should -Invoke Add-CMScriptDeploymentType -Times 1 -Exactly -ParameterFilter { $InstallationBehaviorType -eq $Behavior }
            }
        }

        It 'disables task sequences when any deployment type inherits user context' {
            $script:testManifest | Add-Member -NotePropertyMembers @{
                InstallationBehaviorType = 'InstallForUser'
                DeploymentTypes = @(
                    [pscustomobject]@{ NameSuffix = 'System'; InstallationBehaviorType = 'InstallForSystem' }
                    [pscustomobject]@{ NameSuffix = 'User' }
                )
            }
            Mock Get-CMApplication { $null }
            Mock New-CMApplication { [pscustomobject]@{ CI_ID = 4321 } }
            Mock Add-CMScriptDeploymentType { }
            Mock Remove-CMApplicationRevisionHistoryByCIId { }
            New-MECMApplicationFromManifest -Manifest $script:testManifest -SiteCode MCM -NetworkContentPath '\\server\share\Test' | Out-Null
            Should -Invoke New-CMApplication -Times 1 -Exactly -ParameterFilter { $AutoInstall -eq $false }
            Should -Invoke Add-CMScriptDeploymentType -Times 1 -Exactly -ParameterFilter { $InstallationBehaviorType -eq 'InstallForUser' }
        }

        It 'repairs existing user application task-sequence setting on <Policy>' -TestCases @(
            @{ Policy = 'Skip'; OldVersion = '1.0'; CreatesDt = 0 }
            @{ Policy = 'Overwrite'; OldVersion = '1.0'; CreatesDt = 1 }
            @{ Policy = 'Overwrite'; OldVersion = '0.9'; CreatesDt = 1 }
        ) {
            param($Policy, $OldVersion, $CreatesDt)
            $script:testManifest | Add-Member -NotePropertyName InstallationBehaviorType -NotePropertyValue 'InstallForUser'
            Mock Get-CMApplication { [pscustomobject]@{ CI_ID = 1234; SoftwareVersion = $OldVersion; AutoInstall = $true } }
            $script:taskSequenceDisabled = $false
            Mock Set-CMApplication { if ($PesterBoundParameters.ContainsKey('AutoInstall') -and $AutoInstall -eq $false) { $script:taskSequenceDisabled = $true } }
            Mock Add-CMScriptDeploymentType { if (-not $script:taskSequenceDisabled) { throw 'AutoInstall must be cleared before adding user DT' } }
            Mock Remove-CMDeploymentType { }
            Mock Set-CMDeploymentType { }
            Mock Remove-CMApplicationRevisionHistoryByCIId { }
            New-MECMApplicationFromManifest -Manifest $script:testManifest -SiteCode MCM -NetworkContentPath '\\server\share\Test' -OnExisting $Policy | Should -Be 1234
            Should -Invoke Set-CMApplication -Times 1 -Exactly -ParameterFilter { $AutoInstall -eq $false }
            Should -Invoke Add-CMScriptDeploymentType -Times $CreatesDt -Exactly
        }

        It 'does not mutate an existing user application when OnExisting is Fail' {
            $script:testManifest | Add-Member -NotePropertyName InstallationBehaviorType -NotePropertyValue 'InstallForUser'
            Mock Set-CMApplication { }
            { New-MECMApplicationFromManifest -Manifest $script:testManifest -SiteCode MCM -NetworkContentPath '\\server\share\Test' -OnExisting Fail } | Should -Throw '*OnExisting=Fail*'
            Should -Invoke Set-CMApplication -Times 0 -Exactly
        }

        It 'leaves an already-disabled existing user application unchanged' {
            $script:testManifest | Add-Member -NotePropertyName InstallationBehaviorType -NotePropertyValue 'InstallForUser'
            Mock Get-CMApplication { [pscustomobject]@{ CI_ID = 1234; SoftwareVersion = '1.0'; AutoInstall = $false } }
            Mock Set-CMApplication { }
            New-MECMApplicationFromManifest -Manifest $script:testManifest -SiteCode MCM -NetworkContentPath '\\server\share\Test' -OnExisting Skip | Should -Be 1234
            Should -Invoke Set-CMApplication -Times 0 -Exactly
        }

        It 'preserves a disabled task-sequence setting on an existing system application' {
            Mock Get-CMApplication { [pscustomobject]@{ CI_ID = 1234; SoftwareVersion = '0.9'; AutoInstall = $false } }
            Mock Set-CMApplication { }
            Mock Add-CMScriptDeploymentType { }
            Mock Remove-CMDeploymentType { }
            Mock Set-CMDeploymentType { }
            Mock Remove-CMApplicationRevisionHistoryByCIId { }
            New-MECMApplicationFromManifest -Manifest $script:testManifest -SiteCode MCM -NetworkContentPath '\\server\share\Test' -OnExisting Overwrite | Should -Be 1234
            Should -Invoke Set-CMApplication -Times 0 -Exactly -ParameterFilter { $null -ne $AutoInstall }
        }

        It 'stops before changing deployment types if disabling task sequences fails' {
            $script:testManifest | Add-Member -NotePropertyName InstallationBehaviorType -NotePropertyValue 'InstallForUser'
            Mock Get-CMApplication { [pscustomobject]@{ CI_ID = 1234; SoftwareVersion = '0.9'; AutoInstall = $true } }
            Mock Set-CMApplication { throw 'Eligibility update failed' }
            Mock Add-CMScriptDeploymentType { }
            Mock Remove-CMDeploymentType { }
            { New-MECMApplicationFromManifest -Manifest $script:testManifest -SiteCode MCM -NetworkContentPath '\\server\share\Test' -OnExisting Overwrite } | Should -Throw '*Eligibility update failed*'
            Should -Invoke Add-CMScriptDeploymentType -Times 0 -Exactly
            Should -Invoke Remove-CMDeploymentType -Times 0 -Exactly
        }

        It 'uses resolved system overrides rather than the base user behavior for all-system variants' {
            $script:testManifest | Add-Member -NotePropertyMembers @{
                InstallationBehaviorType = 'InstallForUser'
                DeploymentTypes = @(
                    [pscustomobject]@{ NameSuffix = 'x64'; InstallationBehaviorType = 'InstallForSystem' }
                    [pscustomobject]@{ NameSuffix = 'arm64'; InstallationBehaviorType = 'InstallForSystem' }
                )
            }
            Mock Get-CMApplication { $null }
            Mock New-CMApplication { [pscustomobject]@{ CI_ID = 4321 } }
            Mock Add-CMScriptDeploymentType { }
            Mock Remove-CMApplicationRevisionHistoryByCIId { }
            New-MECMApplicationFromManifest -Manifest $script:testManifest -SiteCode MCM -NetworkContentPath '\\server\share\Test' | Out-Null
            Should -Invoke New-CMApplication -Times 1 -Exactly -ParameterFilter { $AutoInstall -eq $true }
        }

        It 'returns the existing CI_ID when the matching deployment type exists' {
            $result = New-MECMApplicationFromManifest `
                -Manifest $script:testManifest `
                -SiteCode 'MCM' `
                -NetworkContentPath '\\server\share\Applications\Test'

            $result | Should -Be 1234
        }

        It 'fails closed when an existing app is missing the expected deployment type' {
            Mock Get-CMDeploymentType { @() }

            {
                New-MECMApplicationFromManifest `
                    -Manifest $script:testManifest `
                    -SiteCode 'MCM' `
                    -NetworkContentPath '\\server\share\Applications\Test'
            } | Should -Throw '*missing deployment type*'
        }

        It 'replaces the deployment type when the application version changes' {
            Mock Get-CMApplication { [pscustomobject]@{ CI_ID = 1234; SoftwareVersion = '0.9' } }
            Mock Add-CMScriptDeploymentType { }
            Mock Remove-CMDeploymentType { }
            Mock Set-CMDeploymentType { }
            Mock Set-CMApplication { }
            Mock Remove-CMApplicationRevisionHistoryByCIId { }

            $result = New-MECMApplicationFromManifest `
                -Manifest $script:testManifest `
                -SiteCode 'MCM' `
                -OnExisting Overwrite `
                -NetworkContentPath '\\server\share\Applications\Test'

            $result | Should -Be 1234
            Should -Invoke Add-CMScriptDeploymentType -Times 1 -Exactly -ParameterFilter { $DeploymentTypeName -eq 'Test App - 1.0 (staging)' }
            Should -Invoke Remove-CMDeploymentType -Times 1 -Exactly -ParameterFilter { $DeploymentTypeName -eq 'Test App - 1.0' }
            Should -Invoke Set-CMDeploymentType -Times 1 -Exactly -ParameterFilter { $DeploymentTypeName -eq 'Test App - 1.0 (staging)' -and $NewDeploymentTypeName -eq 'Test App - 1.0' }
            Should -Invoke Set-CMApplication -Times 1 -Exactly -ParameterFilter { $SoftwareVersion -eq '1.0' }
        }

        It 'maps GroupSizes to one grouped clause run with an OR connector' {
            $script:testManifest.Detection = [pscustomobject]@{
                Type       = 'Compound'
                Connector  = 'Or'
                GroupSizes = @(2, 2)
                Clauses    = @(
                    [pscustomobject]@{ Type = 'File'; FilePath = 'C:\a'; FileName = 'f.dll'; PropertyType = 'Existence' },
                    [pscustomobject]@{ Type = 'File'; FilePath = 'C:\b'; FileName = 'f.dll'; PropertyType = 'Existence' },
                    [pscustomobject]@{ Type = 'File'; FilePath = 'C:\c'; FileName = 'f.dll'; PropertyType = 'Existence' },
                    [pscustomobject]@{ Type = 'File'; FilePath = 'C:\d'; FileName = 'f.dll'; PropertyType = 'Existence' }
                )
            }
            Mock Get-CMApplication { $null }
            Mock New-CMApplication { [pscustomobject]@{ CI_ID = 4321 } }
            $script:capturedDtParams = $null
            Mock Add-CMScriptDeploymentType { $script:capturedDtParams = $PesterBoundParameters }
            Mock Remove-CMApplicationRevisionHistoryByCIId { }

            $result = New-MECMApplicationFromManifest `
                -Manifest $script:testManifest `
                -SiteCode 'MCM' `
                -NetworkContentPath '\\server\share\Applications\Test'

            $result | Should -Be 4321
            $clauses = @($script:capturedDtParams.AddDetectionClause)
            $clauses.Count | Should -Be 4
            @($script:capturedDtParams.GroupDetectionClauses) | Should -Be @($clauses[2].Setting.LogicalName, $clauses[3].Setting.LogicalName)
            $connectors = @($script:capturedDtParams.DetectionClauseConnector)
            $connectors.Count | Should -Be 1
            $connectors[0].LogicalName | Should -Be $clauses[2].Setting.LogicalName
            $connectors[0].Connector | Should -Be 'OR'
        }

        It 'sends no clause group when the second run holds one clause (<Sizes>)' -ForEach @(
            @{ Sizes = '1,1'; GroupSizes = @(1, 1); Paths = @('C:\a', 'C:\b') }
            @{ Sizes = '2,1'; GroupSizes = @(2, 1); Paths = @('C:\a', 'C:\b', 'C:\c') }
        ) {
            $script:testManifest.Detection = [pscustomobject]@{
                Type       = 'Compound'
                Connector  = 'Or'
                GroupSizes = $GroupSizes
                Clauses    = @($Paths | ForEach-Object { [pscustomobject]@{ Type = 'File'; FilePath = $_; FileName = 'f.dll'; PropertyType = 'Existence' } })
            }
            Mock Get-CMApplication { $null }
            Mock New-CMApplication { [pscustomobject]@{ CI_ID = 4321 } }
            $script:capturedDtParams = $null
            Mock Add-CMScriptDeploymentType { $script:capturedDtParams = $PesterBoundParameters }
            Mock Remove-CMApplicationRevisionHistoryByCIId { }

            New-MECMApplicationFromManifest `
                -Manifest $script:testManifest `
                -SiteCode 'MCM' `
                -NetworkContentPath '\\server\share\Applications\Test' | Should -Be 4321

            $clauses = @($script:capturedDtParams.AddDetectionClause)
            $clauses.Count | Should -Be $Paths.Count
            $script:capturedDtParams.ContainsKey('GroupDetectionClauses') | Should -BeFalse
            $connectors = @($script:capturedDtParams.DetectionClauseConnector)
            $connectors.Count | Should -Be 1
            $connectors[0].LogicalName | Should -Be $clauses[-1].Setting.LogicalName
            $connectors[0].Connector | Should -Be 'OR'
        }

        It 'rejects GroupSizes that do not sum to the clause count' {
            $script:testManifest.Detection = [pscustomobject]@{
                Type       = 'Compound'
                GroupSizes = @(2, 3)
                Clauses    = @(
                    [pscustomobject]@{ Type = 'File'; FilePath = 'C:\a'; FileName = 'f.dll'; PropertyType = 'Existence' },
                    [pscustomobject]@{ Type = 'File'; FilePath = 'C:\b'; FileName = 'f.dll'; PropertyType = 'Existence' }
                )
            }
            Mock Get-CMApplication { $null }
            Mock New-CMApplication { [pscustomobject]@{ CI_ID = 4321 } }

            {
                New-MECMApplicationFromManifest `
                    -Manifest $script:testManifest `
                    -SiteCode 'MCM' `
                    -NetworkContentPath '\\server\share\Applications\Test'
            } | Should -Throw '*GroupSizes*'
        }

        It 'honors manifest InstallCommandLine and UninstallCommandLine overrides' {
            $script:testManifest | Add-Member -NotePropertyName InstallCommandLine -NotePropertyValue 'Invoke-AppDeployToolkit.exe -DeploymentType Install' -Force
            $script:testManifest | Add-Member -NotePropertyName UninstallCommandLine -NotePropertyValue 'Invoke-AppDeployToolkit.exe -DeploymentType Uninstall' -Force
            Mock Get-CMApplication { $null }
            Mock New-CMApplication { [pscustomobject]@{ CI_ID = 5678 } }
            $script:capturedDtParams = $null
            Mock Add-CMScriptDeploymentType { $script:capturedDtParams = $PesterBoundParameters }
            Mock Remove-CMApplicationRevisionHistoryByCIId { }

            New-MECMApplicationFromManifest `
                -Manifest $script:testManifest `
                -SiteCode 'MCM' `
                -NetworkContentPath '\\server\share\Applications\Test' | Out-Null

            $script:capturedDtParams.InstallCommand | Should -Be 'Invoke-AppDeployToolkit.exe -DeploymentType Install'
            $script:capturedDtParams.UninstallCommand | Should -Be 'Invoke-AppDeployToolkit.exe -DeploymentType Uninstall'
        }

        It 'refuses duplicate existing application names' {
            Mock Get-CMApplication {
                @(
                    [pscustomobject]@{ CI_ID = 1001 },
                    [pscustomobject]@{ CI_ID = 1002 }
                )
            }

            {
                New-MECMApplicationFromManifest `
                    -Manifest $script:testManifest `
                    -SiteCode 'MCM' `
                    -NetworkContentPath '\\server\share\Applications\Test'
            } | Should -Throw '*Multiple existing ConfigMgr applications*'
        }

        It 'fails before ConfigMgr app lookup when network content does not match manifest hashes' {
            $contentPath = Join-Path $TestDrive 'network-content-mismatch'
            New-Item -ItemType Directory -Path $contentPath -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $contentPath 'install.ps1') -Value 'expected' -Encoding ASCII
            $script:testManifest | Add-Member -NotePropertyName FileHashes -NotePropertyValue (Get-StageFileHashes -Root $contentPath) -Force
            Set-Content -LiteralPath (Join-Path $contentPath 'install.ps1') -Value 'tampered' -Encoding ASCII
            Mock Get-CMApplication { throw 'Get-CMApplication should not run after integrity failure.' }

            {
                New-MECMApplicationFromManifest `
                    -Manifest $script:testManifest `
                    -SiteCode 'MCM' `
                    -NetworkContentPath $contentPath
            } | Should -Throw '*Package integrity verification failed*'
        }
    }
}

# ============================================================================
# Packager history helpers
# ============================================================================
# Get-PackagerHistoryPath / Read-PackagerHistory / Save-PackagerHistory /
# Update-PackagerHistory back the on-disk app-history.json that the GUI's
# One Click, Check Latest, Stage, and Package paths all share. The file
# lives under $env:LOCALAPPDATA\AppPackager\ in production. Tests redirect
# LOCALAPPDATA into $TestDrive so they never touch real user state.

Describe 'Packager history helpers' {
    BeforeAll {
        $script:OrigLocalAppData = $env:LOCALAPPDATA
        $env:LOCALAPPDATA = Join-Path $TestDrive 'FakeLocalAppData'
    }
    AfterAll {
        $env:LOCALAPPDATA = $script:OrigLocalAppData
    }
    BeforeEach {
        $path = Join-Path $env:LOCALAPPDATA 'AppPackager\app-history.json'
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }

    Context 'Get-PackagerHistoryPath' {
        It 'returns a path under %LOCALAPPDATA%\AppPackager' {
            $p = Get-PackagerHistoryPath
            $p | Should -BeLike (Join-Path $env:LOCALAPPDATA 'AppPackager\app-history.json')
        }

        It 'creates the parent directory on first call' {
            $parent = Join-Path $env:LOCALAPPDATA 'AppPackager'
            if (Test-Path -LiteralPath $parent) { Remove-Item -LiteralPath $parent -Recurse -Force }
            $null = Get-PackagerHistoryPath
            Test-Path -LiteralPath $parent | Should -BeTrue
        }
    }

    Context 'Read-PackagerHistory' {
        It 'returns an empty hashtable when the file is missing' {
            $h = Read-PackagerHistory
            $h | Should -BeOfType [hashtable]
            $h.Count | Should -Be 0
        }

        It 'returns an empty hashtable when the file is empty' {
            $path = Get-PackagerHistoryPath
            Set-Content -LiteralPath $path -Value '' -Encoding UTF8
            $h = Read-PackagerHistory
            $h.Count | Should -Be 0
        }

        It 'returns an empty hashtable when the file is malformed JSON (no throw)' {
            $path = Get-PackagerHistoryPath
            Set-Content -LiteralPath $path -Value '{ not valid json' -Encoding UTF8
            { Read-PackagerHistory } | Should -Not -Throw
            (Read-PackagerHistory).Count | Should -Be 0
        }

        It 'parses valid JSON into a hashtable keyed by packager base name' {
            $path = Get-PackagerHistoryPath
            $data = @{
                'package-chrome' = @{
                    LastChecked      = '2026-04-19T12:00:00Z'
                    LastKnownVersion = '147.0.7727.102'
                    LastResult       = 'NoChange'
                }
            }
            ($data | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $path -Encoding UTF8
            $h = Read-PackagerHistory
            $h.ContainsKey('package-chrome') | Should -BeTrue
            $entry = $h['package-chrome']
            # ConvertFrom-Json returns PSCustomObject so property access is via dot
            ([string]$entry.LastKnownVersion) | Should -Be '147.0.7727.102'
            ([string]$entry.LastResult)       | Should -Be 'NoChange'
        }
    }

    Context 'Save-PackagerHistory' {
        It 'writes JSON that Read-PackagerHistory can parse back' {
            $h = @{
                'package-7zip' = @{
                    LastChecked      = '2026-04-19T12:00:00Z'
                    LastKnownVersion = '26.00'
                    LastResult       = 'Updated'
                }
            }
            Save-PackagerHistory -History $h
            $path = Get-PackagerHistoryPath
            Test-Path -LiteralPath $path | Should -BeTrue

            $roundtrip = Read-PackagerHistory
            $roundtrip.ContainsKey('package-7zip') | Should -BeTrue
            ([string]$roundtrip['package-7zip'].LastKnownVersion) | Should -Be '26.00'
        }
    }

    Context 'Update-PackagerHistory' {
        It 'creates a new entry with ISO 8601 UTC timestamp for Event Checked' {
            Update-PackagerHistory -PackagerName 'package-foo' -Event Checked -Version '1.2.3' -Result Updated

            $h = Read-PackagerHistory
            $h.ContainsKey('package-foo') | Should -BeTrue
            $entry = $h['package-foo']
            ([string]$entry.LastKnownVersion) | Should -Be '1.2.3'
            ([string]$entry.LastResult)       | Should -Be 'Updated'
            ([string]$entry.LastChecked)      | Should -Match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$'
        }

        It 'updates LastStaged for Event Staged and leaves LastChecked alone' {
            Update-PackagerHistory -PackagerName 'package-bar' -Event Checked -Version '1.0' -Result NoChange
            $afterCheck = (Read-PackagerHistory)['package-bar']
            $origChecked = [string]$afterCheck.LastChecked

            Start-Sleep -Seconds 1
            Update-PackagerHistory -PackagerName 'package-bar' -Event Staged -Version '1.0' -Result Updated

            $afterStage = (Read-PackagerHistory)['package-bar']
            ([string]$afterStage.LastChecked) | Should -Be $origChecked
            ([string]$afterStage.LastStaged)  | Should -Match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$'
        }

        It 'updates LastPackaged for Event Packaged' {
            Update-PackagerHistory -PackagerName 'package-baz' -Event Packaged -Version '2.0' -Result Updated

            $entry = (Read-PackagerHistory)['package-baz']
            ([string]$entry.LastPackaged) | Should -Match '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$'
            ([string]$entry.LastKnownVersion) | Should -Be '2.0'
        }

        It 'omits Version update when -Version is not supplied' {
            Update-PackagerHistory -PackagerName 'package-qux' -Event Checked -Version '1.0' -Result NoChange
            Update-PackagerHistory -PackagerName 'package-qux' -Event Checked  # no -Version this time

            $entry = (Read-PackagerHistory)['package-qux']
            ([string]$entry.LastKnownVersion) | Should -Be '1.0'
        }

        It 'records Result=Failed without clobbering prior LastKnownVersion' {
            Update-PackagerHistory -PackagerName 'package-fail' -Event Checked -Version '3.0' -Result Updated
            Update-PackagerHistory -PackagerName 'package-fail' -Event Checked -Result Failed

            $entry = (Read-PackagerHistory)['package-fail']
            ([string]$entry.LastResult)       | Should -Be 'Failed'
            ([string]$entry.LastKnownVersion) | Should -Be '3.0'
        }

        It 'rejects an invalid Event value via ValidateSet' {
            { Update-PackagerHistory -PackagerName 'package-x' -Event 'NotARealEvent' } | Should -Throw
        }

        It 'rejects an invalid Result value via ValidateSet' {
            { Update-PackagerHistory -PackagerName 'package-x' -Event Checked -Result 'Invalid' } | Should -Throw
        }
    }
}

# ============================================================================
# Auto-distribute (Start-CMContentDistribution) integration
# ============================================================================
# New-MECMApplicationFromManifest reads AppPackager.preferences.json relative
# to its own $PSScriptRoot (one level up from Packagers\) and, when
# ContentDistribution.AutoDistribute is true AND DPGroupName is non-empty,
# calls Start-CMContentDistribution after creating the ConfigMgr Application.
#
# Testing this path cleanly requires either:
#   (a) refactoring the auto-distribute block into its own helper that takes
#       the prefs path as a parameter (then tested in isolation), or
#   (b) mocking the full ConfigMgr cmdlet surface (Connect-CMSite,
#       New-CMApplication, Get-CMApplication, Add-CMScriptDeploymentType,
#       Remove-CMApplicationRevisionHistoryByCIId, Start-CMContentDistribution)
#       and writing a temp AppPackager.preferences.json at the real repo root.
#
# Option (a) is the right design; option (b) would clobber dev prefs at repo
# root during test runs. Leaving as a TODO until the refactor lands.

# ============================================================================
# Intune Win32 content prep
# ============================================================================

Describe 'New-IntuneWinPackage' {
    BeforeEach {
        $script:iwRoot = Join-Path $TestDrive 'iw'
        $script:iwContent = Join-Path $script:iwRoot 'content\1.0.0'
        New-Item -ItemType Directory -Path $script:iwContent -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:iwContent 'install.bat') -Value '@echo off' -Encoding Ascii
        $script:iwFakeTool = Join-Path $script:iwRoot 'IntuneWinAppUtil.exe'
        Set-Content -LiteralPath $script:iwFakeTool -Value 'not a real tool' -Encoding Ascii
    }

    It 'throws when the tool path does not exist' {
        { New-IntuneWinPackage -ToolPath (Join-Path $script:iwRoot 'missing.exe') -ContentFolder $script:iwContent -SetupFile 'install.bat' -OutputFolder $script:iwRoot } |
            Should -Throw '*IntuneWinAppUtil.exe not found*'
    }

    It 'throws when the content folder does not exist' {
        { New-IntuneWinPackage -ToolPath $script:iwFakeTool -ContentFolder (Join-Path $script:iwRoot 'nope') -SetupFile 'install.bat' -OutputFolder $script:iwRoot } |
            Should -Throw '*Content folder not found*'
    }

    It 'throws when the setup file is missing from the content folder' {
        { New-IntuneWinPackage -ToolPath $script:iwFakeTool -ContentFolder $script:iwContent -SetupFile 'setup.exe' -OutputFolder $script:iwRoot } |
            Should -Throw '*Setup file not found*'
    }

    It 'refuses an output folder equal to the content folder' {
        { New-IntuneWinPackage -ToolPath $script:iwFakeTool -ContentFolder $script:iwContent -SetupFile 'install.bat' -OutputFolder $script:iwContent } |
            Should -Throw '*must differ from ContentFolder*'
    }
}

Describe 'Install-IntuneWinAppUtil' {
    InModuleScope AppPackagerCommon {
        It 'discards a download whose signature status is not Valid' {
            Mock Write-Log { }
            Mock Invoke-DownloadWithRetry { Set-Content -LiteralPath $OutFile -Value 'x' -Encoding Ascii }
            Mock Get-AuthenticodeSignature { [pscustomobject]@{ Status = 'NotSigned'; SignerCertificate = $null } }

            $dest = Join-Path $TestDrive 'tools-notsigned'
            { Install-IntuneWinAppUtil -DestinationFolder $dest } | Should -Throw '*expected ''Valid''*'

            Test-Path -LiteralPath (Join-Path $dest 'IntuneWinAppUtil.exe') | Should -BeFalse
            @(Get-ChildItem -LiteralPath $dest -Filter '*.tmp' -ErrorAction SilentlyContinue).Count | Should -Be 0
        }

        It 'discards a validly signed download from a non-Microsoft signer' {
            Mock Write-Log { }
            Mock Invoke-DownloadWithRetry { Set-Content -LiteralPath $OutFile -Value 'x' -Encoding Ascii }
            Mock Get-AuthenticodeSignature {
                [pscustomobject]@{
                    Status            = 'Valid'
                    SignerCertificate = [pscustomobject]@{ Subject = 'CN=Example Publisher, O=Example Corp, C=US' }
                }
            }

            $dest = Join-Path $TestDrive 'tools-wrongsigner'
            { Install-IntuneWinAppUtil -DestinationFolder $dest } | Should -Throw '*Microsoft Corporation*'

            Test-Path -LiteralPath (Join-Path $dest 'IntuneWinAppUtil.exe') | Should -BeFalse
            @(Get-ChildItem -LiteralPath $dest -Filter '*.tmp' -ErrorAction SilentlyContinue).Count | Should -Be 0
        }

        It 'installs the file when the signature is Valid and Microsoft-signed' {
            Mock Write-Log { }
            Mock Invoke-DownloadWithRetry { Set-Content -LiteralPath $OutFile -Value 'x' -Encoding Ascii }
            Mock Get-AuthenticodeSignature {
                [pscustomobject]@{
                    Status            = 'Valid'
                    SignerCertificate = [pscustomobject]@{ Subject = 'CN=Microsoft Corporation, O=Microsoft Corporation, L=Redmond, S=Washington, C=US' }
                }
            }

            $dest = Join-Path $TestDrive 'tools-ok'
            $path = Install-IntuneWinAppUtil -DestinationFolder $dest

            $path | Should -Be (Join-Path $dest 'IntuneWinAppUtil.exe')
            Test-Path -LiteralPath $path | Should -BeTrue
            @(Get-ChildItem -LiteralPath $dest -Filter '*.tmp' -ErrorAction SilentlyContinue).Count | Should -Be 0
        }
    }
}

Describe 'Get-InstallerAnalysis' {
    BeforeAll {
        Import-Module "$PSScriptRoot\..\Lib\InstallerAnalysisCommon\InstallerAnalysisCommon.psd1" -Global -DisableNameChecking -Force

        $script:nsisExe = Join-Path $TestDrive 'FakeApp-Setup.exe'
        $bytes = [byte[]](0x4D, 0x5A) + (, [byte]0 * 200) +
            [System.Text.Encoding]::ASCII.GetBytes('NullsoftInst') + (, [byte]0 * 200)
        [System.IO.File]::WriteAllBytes($script:nsisExe, $bytes)
    }

    It 'throws when the file does not exist' {
        { Get-InstallerAnalysis -Path (Join-Path $TestDrive 'missing.msi') } | Should -Throw '*not found*'
    }

    It 'classifies a Nullsoft-marker EXE as predicted NSIS with /S switches' {
        $a = Get-InstallerAnalysis -Path $script:nsisExe
        $a.InstallerType | Should -Be 'NSIS'
        $a.Confidence | Should -Be 'Predicted'
        $a.InstallArgs | Should -Be '/S'
        $a.InstallCommand | Should -Match 'FakeApp-Setup\.exe'
        $a.FileName | Should -Be 'FakeApp-Setup.exe'
    }

    Context 'NSIS header decoded from a makensis build' {
        BeforeAll {
            $script:makensis = 'C:\Program Files (x86)\NSIS\makensis.exe'
            $script:haveMakensis = Test-Path -LiteralPath $script:makensis
            if ($script:haveMakensis) {
                $nsi = Join-Path $TestDrive 'peruser.nsi'
                Set-Content -LiteralPath (Join-Path $TestDrive 'payload.txt') -Value 'payload'
                $script = @(
                    'Unicode true'
                    'SetCompressor /SOLID lzma'
                    'RequestExecutionLevel user'
                    'Name "Drop App"'
                    'OutFile "${OUTFILE}"'
                    'InstallDir "$LOCALAPPDATA\DropApp"'
                    'VIProductVersion "4.5.6.0"'
                    'VIAddVersionKey "ProductName" "Drop App"'
                    'VIAddVersionKey "ProductVersion" "4.5.6"'
                    'VIAddVersionKey "FileVersion" "4.5.6"'
                    'VIAddVersionKey "CompanyName" "Drop Co"'
                    'Section "Main"'
                    '  SetOutPath "$INSTDIR"'
                    '  File "payload.txt"'
                    '  WriteUninstaller "$INSTDIR\Uninstall.exe"'
                    '  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\DropApp" "DisplayName" "Drop App"'
                    '  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\DropApp" "DisplayVersion" "4.5.6"'
                    '  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\DropApp" "Publisher" "Drop Co"'
                    '  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\DropApp" "UninstallString" ''"$INSTDIR\Uninstall.exe"'''
                    'SectionEnd'
                    'Section "Uninstall"'
                    '  Delete "$INSTDIR\payload.txt"'
                    'SectionEnd'
                )
                Set-Content -LiteralPath $nsi -Value $script -Encoding ASCII
                $script:perUserExe = Join-Path $TestDrive 'DropApp-Setup.exe'
                $null = & $script:makensis /V1 "/DOUTFILE=$script:perUserExe" $nsi 2>&1
                $script:haveMakensis = (Test-Path -LiteralPath $script:perUserExe)
            }
        }

        It 'resolves the per-user uninstall path, HKCU key and user context from the script' {
            if (-not $script:haveMakensis) { Set-ItResult -Skipped -Because 'makensis not installed'; return }
            $a = Get-InstallerAnalysis -Path $script:perUserExe
            $a.InstallerType | Should -Be 'NSIS'
            $a.UninstallCommand | Should -Be '"%LOCALAPPDATA%\DropApp\Uninstall.exe" /S'
            $a.UninstallRegistryKey | Should -Be 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\DropApp'
            $a.UninstallRegistryHive | Should -Be 'HKCU'
            $a.InstallContext | Should -Be 'PerUser'
            $a.InstallDir | Should -Be '%LOCALAPPDATA%\DropApp'
            $a.AppName | Should -Be 'Drop App'
            $a.SoftwareVersion | Should -Be '4.5.6'
            $a.RequestedExecutionLevel | Should -Be 'asInvoker'
        }

        It 'stages the per-user drop as an InstallForUser deployment with HKCU detection' {
            if (-not $script:haveMakensis) { Set-ItResult -Skipped -Because 'makensis not installed'; return }
            $a = Get-InstallerAnalysis -Path $script:perUserExe
            $s = New-AdHocStage -Analysis $a -DownloadRoot (Join-Path $TestDrive 'peruser-stage')
            $m = Read-StageManifest -Path $s.ManifestPath
            $m.InstallationBehaviorType | Should -Be 'InstallForUser'
            $m.LogonRequirementType | Should -Be 'OnlyWhenUserLoggedOn'
            $m.InstallContext | Should -Be 'PerUser'
            $m.Detection.Hive | Should -Be 'CurrentUser'
            $m.Detection.RegistryKeyRelative | Should -Be 'Software\Microsoft\Windows\CurrentVersion\Uninstall\DropApp'
            $uninstall = Get-Content (Join-Path $s.StagedPath 'uninstall.ps1') -Raw
            $uninstall | Should -Match ([regex]::Escape("ExpandEnvironmentVariables('%LOCALAPPDATA%\DropApp\Uninstall.exe')"))
            $uninstall | Should -Match ([regex]::Escape("@('/S')"))
        }
    }

    Context 'Inno Setup header decoded from a synthetic setup-0 block' {
        BeforeAll {
            & (Get-Module InstallerAnalysisCommon) { Initialize-InnoBlockType }

            # Data version 6.7.0 layout of Shared.Struct.pas: 39 strings, 4
            # ANSI strings, 17 counts, then the fixed fields up to Options.
            function script:New-InnoFixture {
                param([string]$Path, [string]$AppId, [string]$DefaultDirName, [int]$Privileges, [string]$Arch64 = 'x64compatible', [int]$Overrides = 0)
                $enc = [System.Text.Encoding]::Unicode
                $rec = New-Object System.Collections.Generic.List[byte]
                $addStr = { param([string]$s, [switch]$Ansi) $b = if ($Ansi) { [System.Text.Encoding]::Default.GetBytes($s) } else { $enc.GetBytes($s) }; $rec.AddRange([byte[]][BitConverter]::GetBytes([int32]$b.Length)); if ($b.Length -gt 0) { $rec.AddRange([byte[]]$b) } }
                $zeros = { param($n) $rec.AddRange([byte[]](New-Object byte[] $n)) }
                foreach ($s in @('Drop Inno', '', $AppId, '', 'Drop Co', '', '', '', '', '7.8.9', $DefaultDirName, 'Drop Inno', 'setup', '', '', '', '', '', '', '', '', '', '', '', 'yes', 'yes', '', '', 'no', 'no', 'x64compatible', $Arch64, '', '', 'yes', 'yes', 'yes', 'yes', 'yes')) { & $addStr $s }
                foreach ($s in @('', '', '', '')) { & $addStr $s -Ansi }
                $rec.AddRange([byte[]][BitConverter]::GetBytes([int32]1)); & $zeros 64
                $rec.AddRange([byte[]]@(0, 0, 0, 0, 0, 0, 1, 6, 0, 0)); & $zeros 10
                & $zeros 8; $rec.Add([byte]2); $rec.Add([byte]1); & $zeros 26; $rec.Add([byte]0)
                & $zeros 8; $rec.AddRange([byte[]][BitConverter]::GetBytes([int32]1))
                $rec.AddRange([byte[]]@(1, 0, $Privileges, $Overrides, 2, 0, 3, 0, 0)); & $zeros 16
                $record = $rec.ToArray()

                $crc = { param($bytes, $off, $len) [InstallerAnalysis.InnoBlock]::Crc32($bytes, $off, $len) }
                $stub = New-Object byte[] 1024
                $stub[0] = 0x4D; $stub[1] = 0x5A
                $table = New-Object System.Collections.Generic.List[byte]
                $table.AddRange([byte[]]@(0x72,0x44,0x6C,0x50,0x74,0x53,0xCD,0xE6,0xD7,0x7B,0x0B,0x2A))
                $table.AddRange([byte[]][BitConverter]::GetBytes([uint32]2))
                foreach ($v in @([int64]0, [int64]0)) { $table.AddRange([byte[]][BitConverter]::GetBytes($v)) }
                $table.AddRange([byte[]][BitConverter]::GetBytes([uint32]0)); $table.AddRange([byte[]][BitConverter]::GetBytes([int32]0))
                $table.AddRange([byte[]][BitConverter]::GetBytes([int64]1024)); $table.AddRange([byte[]][BitConverter]::GetBytes([int64]0))
                $table.AddRange([byte[]][BitConverter]::GetBytes([uint32]0))
                $tb = $table.ToArray(); $table.AddRange([byte[]][BitConverter]::GetBytes([uint32](& $crc $tb 0 60)))
                [Array]::Copy($table.ToArray(), 0, $stub, 64, 64)

                $out = New-Object System.Collections.Generic.List[byte]
                $out.AddRange([byte[]]$stub)
                $id = New-Object byte[] 64
                $idBytes = [System.Text.Encoding]::ASCII.GetBytes('Inno Setup Setup Data (6.7.0)'); [Array]::Copy($idBytes, 0, $id, 0, $idBytes.Length)
                $out.AddRange([byte[]]$id)
                $encHeader = New-Object byte[] 49
                $out.AddRange([byte[]][BitConverter]::GetBytes([uint32](& $crc $encHeader 0 49))); $out.AddRange([byte[]]$encHeader)
                $chunks = New-Object System.Collections.Generic.List[byte]
                $pos = 0
                while ($pos -lt $record.Length) {
                    $len = [Math]::Min(4096, $record.Length - $pos)
                    $chunks.AddRange([byte[]][BitConverter]::GetBytes([uint32](& $crc $record $pos $len)))
                    for ($i = 0; $i -lt $len; $i++) { $chunks.Add($record[$pos + $i]) }
                    $pos += $len
                }
                $stored = $chunks.ToArray()
                $bh = New-Object System.Collections.Generic.List[byte]
                $bh.AddRange([byte[]][BitConverter]::GetBytes([int64]$stored.Length)); $bh.Add([byte]0)
                $bhBytes = $bh.ToArray()
                $out.AddRange([byte[]][BitConverter]::GetBytes([uint32](& $crc $bhBytes 0 $bhBytes.Length))); $out.AddRange([byte[]]$bhBytes); $out.AddRange([byte[]]$stored)
                [System.IO.File]::WriteAllBytes($Path, $out.ToArray())
                return $Path
            }
            $script:innoMachine = New-InnoFixture -Path (Join-Path $TestDrive 'DropInno-Setup.exe') -AppId '{{6D2A0E1C-1B4F-4C2B-9E0B-6C1D5B2A7F10}' -DefaultDirName '{autopf}\Drop Inno' -Privileges 2
            $script:innoUser = New-InnoFixture -Path (Join-Path $TestDrive 'DropInnoUser-Setup.exe') -AppId 'DropInnoUser' -DefaultDirName '{autopf}\Drop Inno' -Privileges 3
        }

        It 'takes the AppId key, 64-bit view, Program Files folder and unins000.exe from the header' {
            $a = Get-InstallerAnalysis -Path $script:innoMachine
            $a.InstallerType | Should -Be 'InnoSetup'
            $a.AppName | Should -Be 'Drop Inno'
            $a.SoftwareVersion | Should -Be '7.8.9'
            $a.Publisher | Should -Be 'Drop Co'
            $a.UninstallRegistryKey | Should -Be 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{6D2A0E1C-1B4F-4C2B-9E0B-6C1D5B2A7F10}_is1'
            $a.UninstallRegistryHive | Should -Be 'HKLM'
            $a.RegistryView | Should -Be '64'
            $a.InstallContext | Should -Be 'PerMachine'
            $a.InstallDir | Should -Be '%ProgramFiles%\Drop Inno'
            $a.UninstallCommand | Should -Be '"%ProgramFiles%\Drop Inno\unins000.exe" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART'
            $a.InstallArgs | Should -Be '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-'
        }

        It 'stages the per-machine drop with the GUID key as the detection target' {
            $a = Get-InstallerAnalysis -Path $script:innoMachine
            $s = New-AdHocStage -Analysis $a -DownloadRoot (Join-Path $TestDrive 'inno-stage')
            $m = Read-StageManifest -Path $s.ManifestPath
            $m.InstallationBehaviorType | Should -Not -Be 'InstallForUser'
            $m.InstallContext | Should -Be 'PerMachine'
            $m.Detection.RegistryKeyRelative | Should -Be 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{6D2A0E1C-1B4F-4C2B-9E0B-6C1D5B2A7F10}_is1'
            $m.Detection.Is64Bit | Should -BeTrue
            $uninstall = Get-Content (Join-Path $s.StagedPath 'uninstall.ps1') -Raw
            $uninstall | Should -Match ([regex]::Escape("ExpandEnvironmentVariables('%ProgramFiles%\Drop Inno\unins000.exe')"))
        }

        It 'maps PrivilegesRequired=lowest to a user-context deployment with HKCU detection' {
            $a = Get-InstallerAnalysis -Path $script:innoUser
            $a.InstallContext | Should -Be 'PerUser'
            $a.UninstallRegistryKey | Should -Be 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\DropInnoUser_is1'
            $a.InstallDir | Should -Be '%LOCALAPPDATA%\Programs\Drop Inno'
            $s = New-AdHocStage -Analysis $a -DownloadRoot (Join-Path $TestDrive 'inno-user-stage')
            $m = Read-StageManifest -Path $s.ManifestPath
            $m.InstallationBehaviorType | Should -Be 'InstallForUser'
            $m.Detection.Hive | Should -Be 'CurrentUser'
        }
    }

    Context 'Install mode switch for a switchable installer' {
        BeforeAll {
            $script:innoSwitchable = New-InnoFixture -Path (Join-Path $TestDrive 'DropTree-Setup.exe') -AppId 'DropTree' -DefaultDirName '{autopf}\Drop Tree' -Privileges 3 -Overrides 1
        }

        It 'reports both modes and describes the installer default' {
            $a = Get-InstallerAnalysis -Path $script:innoSwitchable
            $a.InstallMode | Should -Be 'CurrentUser'
            @($a.InstallModes) | Should -Be @('CurrentUser', 'AllUsers')
            $a.InstallContext | Should -Be 'PerUser'
            $a.InstallArgs | Should -Be '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-'
            $a.UninstallRegistryKey | Should -Be 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\DropTree_is1'
        }

        It 'switches every mode-dependent field together and stages a system deployment from the all-users branch' {
            $a = Get-InstallerAnalysis -Path $script:innoSwitchable
            $s = Set-InstallerAnalysisMode -Analysis $a -Mode AllUsers
            $s.InstallMode | Should -Be 'AllUsers'
            $s.InstallArgs | Should -Be '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /ALLUSERS'
            $s.InstallCommand | Should -Be '"DropTree-Setup.exe" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP- /ALLUSERS'
            $s.UninstallCommand | Should -Be '"%ProgramFiles%\Drop Tree\unins000.exe" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART'
            $s.UninstallRegistryKey | Should -Be 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\DropTree_is1'
            $s.UninstallRegistryHive | Should -Be 'HKLM'
            $s.RegistryView | Should -Be '64'
            $s.InstallDir | Should -Be '%ProgramFiles%\Drop Tree'
            $s.InstallContext | Should -Be 'PerMachine'
            # the original analysis is untouched
            $a.InstallContext | Should -Be 'PerUser'

            $stage = New-AdHocStage -Analysis $s -DownloadRoot (Join-Path $TestDrive 'mode-stage')
            $m = Read-StageManifest -Path $stage.ManifestPath
            $m.InstallContext | Should -Be 'PerMachine'
            $m.InstallationBehaviorType | Should -Not -Be 'InstallForUser'
            $m.Detection.RegistryKeyRelative | Should -Be 'Software\Microsoft\Windows\CurrentVersion\Uninstall\DropTree_is1'
            $m.Detection.Hive | Should -Not -Be 'CurrentUser'
            $install = Get-Content (Join-Path $stage.StagedPath 'install.ps1') -Raw
            $install | Should -Match '/ALLUSERS'
            $uninstall = Get-Content (Join-Path $stage.StagedPath 'uninstall.ps1') -Raw
            $uninstall | Should -Match ([regex]::Escape('%ProgramFiles%\Drop Tree\unins000.exe'))
        }

        It 'refuses a mode the installer does not offer' {
            $single = New-InnoFixture -Path (Join-Path $TestDrive 'DropSingle-Setup.exe') -AppId 'DropSingle' -DefaultDirName '{autopf}\Drop Single' -Privileges 2
            $a = Get-InstallerAnalysis -Path $single
            @($a.InstallModes) | Should -Be @('AllUsers')
            { Set-InstallerAnalysisMode -Analysis $a -Mode CurrentUser } | Should -Throw '*offers install mode*'
        }
    }

    Context 'Manifest-only context inference' {
        It 'predicts per-user only for an EXE whose format carries no context' {
            $a = Get-InstallerAnalysis -Path (Join-Path $env:SystemRoot 'System32\notepad.exe')
            $a.InstallerType | Should -Be 'Unknown'
            $a.RequestedExecutionLevel | Should -Be 'asInvoker'
            $a.InstallContext | Should -Be 'PerUser'
        }
    }

    Context 'MSI extraction (mocked reader functions)' {
        BeforeAll {
            Mock Get-InstallerType { 'MSI' } -ModuleName AppPackagerCommon
            Mock Get-InstallerFileInfo {
                [pscustomobject]@{ Architecture = 'N/A (see MSI Summary)' }
            } -ModuleName AppPackagerCommon
            Mock Get-MsiProperties {
                @{ ProductCode = '{11111111-2222-3333-4444-555555555555}'; ProductName = 'Widget'; Manufacturer = 'Acme'; ProductVersion = '3.1.0' }
            } -ModuleName AppPackagerCommon
            Mock Get-MsiSummaryInfo { [pscustomobject]@{ Architecture = 'x64' } } -ModuleName AppPackagerCommon
            Mock Get-PackageMetadataFor { $null } -ModuleName AppPackagerCommon
            Mock Get-SilentSwitches {
                [pscustomobject]@{ Install = 'msiexec /i "f.msi" /qn'; Uninstall = 'msiexec /x {11111111-2222-3333-4444-555555555555} /qn'; Notes = '' }
            } -ModuleName AppPackagerCommon
            Mock Get-DeploymentFields {
                [pscustomobject]@{ DisplayName = 'Widget'; DisplayVersion = '3.1.0'; Vendor = 'Acme'; SilentUninstallString = ''; UninstallRegistryKey = '' }
            } -ModuleName AppPackagerCommon

            $script:fakeMsi = Join-Path $TestDrive 'widget.msi'
            Set-Content -LiteralPath $script:fakeMsi -Value 'not a real msi' -Encoding ASCII
        }

        It 'returns authoritative identity from the MSI tables' {
            $a = Get-InstallerAnalysis -Path $script:fakeMsi
            $a.Confidence | Should -Be 'Authoritative'
            $a.AppName | Should -Be 'Widget'
            $a.Publisher | Should -Be 'Acme'
            $a.SoftwareVersion | Should -Be '3.1.0'
            $a.ProductCode | Should -Be '{11111111-2222-3333-4444-555555555555}'
            $a.InstallArgs | Should -Be '/qn /norestart'
        }

        It 'prefers the MSI summary architecture over file info' {
            (Get-InstallerAnalysis -Path $script:fakeMsi).Architecture | Should -Be 'x64'
        }
    }
}

Describe 'New-AdHocStage' {
    BeforeAll {
        $script:root = Join-Path $TestDrive 'adhoc'

        $script:msiSource = Join-Path $TestDrive 'drop.msi'
        Set-Content -LiteralPath $script:msiSource -Value 'msi payload' -Encoding ASCII
        $script:msiAnalysis = [pscustomobject]@{
            Path = $script:msiSource; FileName = 'drop.msi'; InstallerType = 'MSI'
            AppName = 'Widget'; Publisher = 'Acme'; SoftwareVersion = '3.1.0'
            ProductCode = '{11111111-2222-3333-4444-555555555555}'
            InstallArgs = '/qn /norestart'; UninstallArgs = '/qn /norestart'
            InstallCommand = ''; UninstallCommand = ''; UninstallRegistryKey = ''
            Architecture = 'x64'; Confidence = 'Authoritative'
        }

        $script:exeSource = Join-Path $TestDrive 'drop-setup.exe'
        Set-Content -LiteralPath $script:exeSource -Value 'exe payload' -Encoding ASCII
        $script:exeAnalysis = [pscustomobject]@{
            Path = $script:exeSource; FileName = 'drop-setup.exe'; InstallerType = 'NSIS'
            AppName = 'Gadget'; Publisher = 'Bmce'; SoftwareVersion = '2.0'
            ProductCode = ''
            InstallArgs = '/S'; UninstallArgs = ''
            InstallCommand = '"drop-setup.exe" /S'; UninstallCommand = '"uninstall.exe" /S'
            UninstallRegistryKey = 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Gadget_is1'
            Architecture = 'x86'; Confidence = 'Predicted'
        }
    }

    It 'stages an MSI with ARP ProductCode detection' {
        $s = New-AdHocStage -Analysis $script:msiAnalysis -DownloadRoot $script:root
        Test-Path (Join-Path $s.StagedPath 'drop.msi') | Should -BeTrue
        foreach ($w in 'install.bat', 'install.ps1', 'uninstall.bat', 'uninstall.ps1') {
            Test-Path (Join-Path $s.StagedPath $w) | Should -BeTrue
        }
        $m = Read-StageManifest -Path $s.ManifestPath
        $m.InstallerType | Should -Be 'MSI'
        $m.Detection.RegistryKeyRelative | Should -Be 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{11111111-2222-3333-4444-555555555555}'
        $m.Detection.Is64Bit | Should -BeTrue
        $m.Detection.ValueName | Should -Be 'DisplayVersion'
    }

    It 'stages an EXE with quoted wrapper args and a split uninstall command' {
        $s = New-AdHocStage -Analysis $script:exeAnalysis -DownloadRoot $script:root
        $install = Get-Content (Join-Path $s.StagedPath 'install.ps1') -Raw
        $install | Should -Match ([regex]::Escape("@('/S')"))
        $uninstall = Get-Content (Join-Path $s.StagedPath 'uninstall.ps1') -Raw
        $uninstall | Should -Match ([regex]::Escape("ExpandEnvironmentVariables('uninstall.exe')"))
        $uninstall | Should -Match ([regex]::Escape("-FilePath $uninstallPath"))
        $uninstall | Should -Match ([regex]::Escape("@('/S')"))
        foreach ($w in 'install.ps1', 'uninstall.ps1') {
            $parseErrors = $null
            [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $s.StagedPath $w), [ref]$null, [ref]$parseErrors) | Out-Null
            $parseErrors.Count | Should -Be 0
        }
    }

    It 'normalizes manifest type to EXE, keeps the engine, strips WOW6432Node into Is64Bit=false' {
        $s = New-AdHocStage -Analysis $script:exeAnalysis -DownloadRoot (Join-Path $script:root 'x')
        $m = Read-StageManifest -Path $s.ManifestPath
        $m.InstallerType | Should -Be 'EXE'
        $m.DetectedEngine | Should -Be 'NSIS'
        $m.Detection.RegistryKeyRelative | Should -Be 'Software\Microsoft\Windows\CurrentVersion\Uninstall\Gadget_is1'
        $m.Detection.Is64Bit | Should -BeFalse
        $m.UninstallCommand | Should -Be '"uninstall.exe" /S'
    }

    It 'sanitizes identity fields into safe folder segments' {
        $dirty = $script:msiAnalysis.PSObject.Copy()
        $s = New-AdHocStage -Analysis $dirty -DownloadRoot (Join-Path $script:root 'y') `
            -AppName 'Bad:Name?' -Publisher 'A/B\C'
        $s.AppFolder | Should -Be 'BadName'
        $s.VendorFolder | Should -Be 'ABC'
        Test-Path $s.StagedPath | Should -BeTrue
    }

    It 'refuses to stage without a version' {
        $noVer = $script:msiAnalysis.PSObject.Copy()
        $noVer.SoftwareVersion = ''
        { New-AdHocStage -Analysis $noVer -DownloadRoot $script:root } | Should -Throw '*version*'
    }

    It 'writes a per-user manifest from an HKCU-registered analysis and keeps a 32-bit view out of Is64Bit' {
        $perUser = $script:exeAnalysis.PSObject.Copy()
        $perUser | Add-Member -NotePropertyName InstallContext -NotePropertyValue 'PerUser' -Force
        $perUser | Add-Member -NotePropertyName RegistryView -NotePropertyValue '' -Force
        $perUser.UninstallRegistryKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Gadget'
        $perUser.UninstallCommand = '"%LOCALAPPDATA%\Gadget\Uninstall.exe" /S'
        $s = New-AdHocStage -Analysis $perUser -DownloadRoot (Join-Path $script:root 'peruser')
        $m = Read-StageManifest -Path $s.ManifestPath
        $m.InstallationBehaviorType | Should -Be 'InstallForUser'
        $m.LogonRequirementType | Should -Be 'OnlyWhenUserLoggedOn'
        $m.Detection.Hive | Should -Be 'CurrentUser'
        $m.Detection.RegistryKeyRelative | Should -Be 'Software\Microsoft\Windows\CurrentVersion\Uninstall\Gadget'
        $m.Detection.Is64Bit | Should -BeFalse
        (Get-Content (Join-Path $s.StagedPath 'uninstall.ps1') -Raw) | Should -Match ([regex]::Escape("ExpandEnvironmentVariables('%LOCALAPPDATA%\Gadget\Uninstall.exe')"))

        $view32 = $script:exeAnalysis.PSObject.Copy()
        $view32 | Add-Member -NotePropertyName RegistryView -NotePropertyValue '32' -Force
        $view32.UninstallRegistryKey = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Gadget'
        $view32.Architecture = 'x64'
        $s2 = New-AdHocStage -Analysis $view32 -DownloadRoot (Join-Path $script:root 'view32')
        $m2 = Read-StageManifest -Path $s2.ManifestPath
        $m2.Detection.Is64Bit | Should -BeFalse
        $m2.Detection.Hive | Should -Be 'LocalMachine'
        $m2.PSObject.Properties['InstallationBehaviorType'] | Should -BeNullOrEmpty
    }

    It 'refuses a non-MSI stage without install arguments' {
        $noArgs = $script:exeAnalysis.PSObject.Copy()
        $noArgs.InstallArgs = ''
        { New-AdHocStage -Analysis $noArgs -DownloadRoot $script:root } | Should -Throw '*install arguments*'
    }
}

Describe 'Velopack drop' {
    BeforeAll {
        $script:veloRoot = Join-Path $TestDrive 'velo'
        $payloadDir = Join-Path $script:veloRoot 'pkg'
        New-Item -ItemType Directory -Path $payloadDir -Force | Out-Null
        $nuspec = @'
<?xml version="1.0" encoding="utf-8"?>
<package xmlns="http://schemas.microsoft.com/packaging/2013/05/nuspec.xsd">
  <metadata>
    <id>WidgetVelo</id>
    <version>4.2.1-build.7</version>
    <title>Widget Velo</title>
    <authors>Acme</authors>
    <machineArchitecture>x64</machineArchitecture>
    <mainExe>WidgetVelo.exe</mainExe>
  </metadata>
</package>
'@
        Set-Content -LiteralPath (Join-Path $payloadDir 'WidgetVelo.nuspec') -Value $nuspec -Encoding UTF8
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zipPath = Join-Path $script:veloRoot 'pkg.zip'
        [IO.Compression.ZipFile]::CreateFromDirectory($payloadDir, $zipPath)
        $zipBytes = [IO.File]::ReadAllBytes($zipPath)
        # setup.rs bundle placeholder: offset and length precede the signature.
        $signature = [byte[]]@(
            0x94,0xf0,0xb1,0x7b,0x68,0x93,0xe0,0x29,0x37,0xeb,0x34,0xef,0x53,0xaa,0xe7,0xd4,
            0x2b,0x54,0xf5,0x70,0x7e,0xf5,0xd6,0xf5,0x78,0x54,0x98,0x3e,0x5e,0x94,0xed,0x7d
        )
        $stub = New-Object byte[] 128
        $stub[0] = 0x4D; $stub[1] = 0x5A
        $bundleOffset = [int64]($stub.Length + 48)
        $bytes = New-Object System.Collections.Generic.List[byte]
        $bytes.AddRange($stub)
        $bytes.AddRange([BitConverter]::GetBytes($bundleOffset))
        $bytes.AddRange([BitConverter]::GetBytes([int64]$zipBytes.Length))
        $bytes.AddRange($signature)
        $bytes.AddRange($zipBytes)
        $script:veloSetup = Join-Path $script:veloRoot 'WidgetVeloSetup.exe'
        [IO.File]::WriteAllBytes($script:veloSetup, $bytes.ToArray())
        $script:veloAnalysis = Get-InstallerAnalysis -Path $script:veloSetup
    }

    It 'reads the nuspec identity, per-user context and application architecture' {
        $a = $script:veloAnalysis
        $a.InstallerType | Should -Be 'Velopack'
        $a.AppName | Should -Be 'Widget Velo'
        $a.SoftwareVersion | Should -Be '4.2.1'
        $a.InstallContext | Should -Be 'PerUser'
        $a.InstallArgs | Should -Be '--silent'
        $a.Architecture | Should -Be 'x64'
        $a.UninstallRegistryKey | Should -Be 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\WidgetVelo'
        $a.UninstallCommand | Should -Be '"%LOCALAPPDATA%\WidgetVelo\Update.exe" --uninstall --silent'
    }

    It 'stages HKCU detection, a user deployment type and an Update.exe uninstall' {
        $s = New-AdHocStage -Analysis $script:veloAnalysis -DownloadRoot (Join-Path $script:veloRoot 'stage')
        $m = Read-StageManifest -Path $s.ManifestPath
        $m.InstallerType | Should -Be 'EXE'
        $m.DetectedEngine | Should -Be 'Velopack'
        $m.InstallContext | Should -Be 'PerUser'
        $m.InstallationBehaviorType | Should -Be 'InstallForUser'
        $m.LogonRequirementType | Should -Be 'OnlyWhenUserLoggedOn'
        $m.InstallArgs | Should -Be '--silent'
        $m.Detection.Type | Should -Be 'RegistryKeyValue'
        $m.Detection.Hive | Should -Be 'CurrentUser'
        $m.Detection.RegistryKeyRelative | Should -Be 'Software\Microsoft\Windows\CurrentVersion\Uninstall\WidgetVelo'
        $m.Detection.DisplayVersion | Should -Be '4.2.1'
        $m.Detection.Is64Bit | Should -BeTrue
        (Get-Content (Join-Path $s.StagedPath 'install.ps1') -Raw) | Should -Match ([regex]::Escape("@('--silent')"))
        $uninstall = Get-Content (Join-Path $s.StagedPath 'uninstall.ps1') -Raw
        $uninstall | Should -Match ([regex]::Escape("ExpandEnvironmentVariables('%LOCALAPPDATA%\WidgetVelo\Update.exe')"))
        $uninstall | Should -Match ([regex]::Escape("@('--uninstall', '--silent')"))
    }
}

Describe 'Invoke-AdHocPackage' {
    BeforeAll {
        $source = Join-Path $TestDrive 'pkg.msi'
        Set-Content -LiteralPath $source -Value 'payload' -Encoding ASCII
        $analysis = [pscustomobject]@{
            Path = $source; FileName = 'pkg.msi'; InstallerType = 'MSI'
            AppName = 'Widget'; Publisher = 'Acme'; SoftwareVersion = '3.1.0'
            ProductCode = '{11111111-2222-3333-4444-555555555555}'
            InstallArgs = '/qn /norestart'; UninstallArgs = '/qn /norestart'
            InstallCommand = ''; UninstallCommand = ''; UninstallRegistryKey = ''
            Architecture = 'x64'; Confidence = 'Authoritative'
        }
        $script:stage = New-AdHocStage -Analysis $analysis -DownloadRoot (Join-Path $TestDrive 'pkgroot')

        $script:share = Join-Path $TestDrive 'share'
        New-Item -ItemType Directory -Force $script:share | Out-Null
    }

    It 'copies content to the network version folder and creates the application' {
        Mock New-MECMApplicationFromManifest {
            [pscustomobject]@{ AppName = $Manifest.AppName; ContentPath = $NetworkContentPath }
        } -ModuleName AppPackagerCommon

        $result = Invoke-AdHocPackage -StagedPath $script:stage.StagedPath `
            -VendorFolder $script:stage.VendorFolder -AppFolder $script:stage.AppFolder `
            -FileServerPath $script:share -SiteCode 'LAB'

        $expected = Join-Path $script:share 'Applications\Acme\Widget\3.1.0'
        Test-Path (Join-Path $expected 'pkg.msi') | Should -BeTrue
        Test-Path (Join-Path $expected 'install.ps1') | Should -BeTrue
        Test-Path (Join-Path $expected 'stage-manifest.json') | Should -BeFalse
        $result.ContentPath | Should -Be $expected
        Should -Invoke New-MECMApplicationFromManifest -Times 1 -Exactly -ModuleName AppPackagerCommon
    }

    It 'throws when the network root is not accessible' {
        { Invoke-AdHocPackage -StagedPath $script:stage.StagedPath `
            -VendorFolder 'Acme' -AppFolder 'Widget' `
            -FileServerPath (Join-Path $TestDrive 'no-such-share') -SiteCode 'LAB' } |
            Should -Throw '*not accessible*'
    }
}

Describe 'New-PackagerFromDrop' {
    BeforeAll {
        $script:repo = Join-Path $TestDrive 'repo'
        New-Item -ItemType Directory -Force (Join-Path $script:repo 'Packagers'), (Join-Path $script:repo 'Samples') | Out-Null
        Copy-Item "$PSScriptRoot\..\Samples\package-template-msi.ps1" (Join-Path $script:repo 'Samples') -Force
        Copy-Item "$PSScriptRoot\..\Samples\package-template-exe.ps1" (Join-Path $script:repo 'Samples') -Force

        $script:msiAnalysis = [pscustomobject]@{
            Path = 'x'; FileName = 'widget-3.1.0.msi'; InstallerType = 'MSI'
            AppName = 'Widget'; Publisher = 'Acme'; SoftwareVersion = '3.1.0'
            ProductCode = '{11111111-2222-3333-4444-555555555555}'
            InstallArgs = '/qn /norestart'; UninstallArgs = ''
            InstallCommand = ''; UninstallCommand = ''; UninstallRegistryKey = ''
            Architecture = 'x64'; Confidence = 'Authoritative'
        }
        $script:exeAnalysis = [pscustomobject]@{
            Path = 'x'; FileName = 'gadget-setup.exe'; InstallerType = 'InnoSetup'
            AppName = 'Gadget'; Publisher = 'Bmce'; SoftwareVersion = '2.0'
            ProductCode = ''
            InstallArgs = '/VERYSILENT /NORESTART'; UninstallArgs = ''
            InstallCommand = ''; UninstallCommand = ''; UninstallRegistryKey = ''
            Architecture = 'x64'; Confidence = 'Predicted'
        }
    }

    It 'fills MSI identity into the template and the result parses' {
        $p = New-PackagerFromDrop -Analysis $script:msiAnalysis -PackagersRoot (Join-Path $script:repo 'Packagers')
        Split-Path -Leaf $p | Should -Be 'package-widget.ps1'
        $c = Get-Content -LiteralPath $p -Raw
        $c | Should -Match 'Vendor: Acme'
        $c | Should -Match 'CMName: Widget'
        $c | Should -Not -Match 'Vendor: TODO'
        $c | Should -Match ([regex]::Escape('$MsiFileName      = "widget-3.1.0.msi"'))
        $parseErrors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($p, [ref]$null, [ref]$parseErrors) | Out-Null
        $parseErrors.Count | Should -Be 0
    }

    It 'fills EXE identity and predicted install args into the template' {
        $p = New-PackagerFromDrop -Analysis $script:exeAnalysis -PackagersRoot (Join-Path $script:repo 'Packagers')
        $c = Get-Content -LiteralPath $p -Raw
        $c | Should -Match ([regex]::Escape('AppName          = "Gadget"'))
        $c | Should -Match ([regex]::Escape('$installArgs   = "/VERYSILENT /NORESTART"'))
        $parseErrors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($p, [ref]$null, [ref]$parseErrors) | Out-Null
        $parseErrors.Count | Should -Be 0
    }

    It 'refuses to overwrite an existing packager' {
        { New-PackagerFromDrop -Analysis $script:msiAnalysis -PackagersRoot (Join-Path $script:repo 'Packagers') } |
            Should -Throw '*already exists*'
    }

    It 'refuses generation without an application name' {
        $anon = $script:exeAnalysis.PSObject.Copy()
        $anon.AppName = ''
        { New-PackagerFromDrop -Analysis $anon -PackagersRoot (Join-Path $script:repo 'Packagers') } |
            Should -Throw '*application name*'
    }
}


Describe 'Ad-hoc intake hardening' {
    BeforeAll {
        $src = Join-Path $TestDrive "Bob's Setup.exe"
        Set-Content -LiteralPath $src -Value 'payload' -Encoding ASCII
        $script:quotedExe = [pscustomobject]@{
            Path = $src; FileName = "Bob's Setup.exe"; InstallerType = 'NSIS'
            AppName = 'Bobs App'; Publisher = 'Bob'; SoftwareVersion = '1.0'
            ProductCode = ''
            InstallArgs = '/S'; UninstallArgs = ''
            InstallCommand = ''; UninstallCommand = "C:\Program Files\Bob's App\uninstall.exe /S"
            UninstallRegistryKey = ''
            Architecture = 'x64'; Confidence = 'Predicted'
        }
    }

    It 'escapes apostrophes in the installer filename and uninstall path so wrappers parse' {
        $s = New-AdHocStage -Analysis $script:quotedExe -DownloadRoot (Join-Path $TestDrive 'q')
        foreach ($w in 'install.ps1', 'uninstall.ps1') {
            $parseErrors = $null
            [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $s.StagedPath $w), [ref]$null, [ref]$parseErrors) | Out-Null
            $parseErrors.Count | Should -Be 0
        }
        (Get-Content (Join-Path $s.StagedPath 'install.ps1') -Raw) | Should -Match ([regex]::Escape("Bob''s Setup.exe"))
    }

    It 'refuses an MSP drop instead of wrapping it as an EXE' {
        $msp = $script:quotedExe.PSObject.Copy()
        $msp.InstallerType = 'MSP'
        { New-AdHocStage -Analysis $msp -DownloadRoot (Join-Path $TestDrive 'q2') } | Should -Throw '*not supported*'
    }

    It 'refuses a non-ASCII installer filename' {
        $uni = $script:quotedExe.PSObject.Copy()
        $uni.FileName = "caf$([char]0xE9)-setup.exe"
        { New-AdHocStage -Analysis $uni -DownloadRoot (Join-Path $TestDrive 'q3') } | Should -Throw '*non-ASCII*'
    }

    It 'refuses identity fields that sanitize to empty path segments' {
        $bad = $script:quotedExe.PSObject.Copy()
        $bad.SoftwareVersion = '***'
        { New-AdHocStage -Analysis $bad -DownloadRoot (Join-Path $TestDrive 'q4') } | Should -Throw '*no usable path characters*'
    }

    Context 'generator neutralizes hostile identity values' {
        BeforeAll {
            $script:genRepo = Join-Path $TestDrive 'hardrepo'
            New-Item -ItemType Directory -Force (Join-Path $script:genRepo 'Packagers'), (Join-Path $script:genRepo 'Samples') | Out-Null
            Copy-Item "$PSScriptRoot\..\Samples\package-template-msi.ps1" (Join-Path $script:genRepo 'Samples') -Force
            Copy-Item "$PSScriptRoot\..\Samples\package-template-exe.ps1" (Join-Path $script:genRepo 'Samples') -Force
        }

        It 'a version resource with embedded quotes and subexpressions cannot inject statements' {
            $hostile = [pscustomobject]@{
                Path = 'x'; FileName = 'evil-setup.exe'; InstallerType = 'InnoSetup'
                AppName = 'Foo" ; Invoke-InjectedMarker ; $x="'
                Publisher = 'Bar$(Invoke-OtherMarker)'
                SoftwareVersion = '1.0'; ProductCode = ''
                InstallArgs = '/VERYSILENT'; UninstallArgs = ''
                InstallCommand = ''; UninstallCommand = ''; UninstallRegistryKey = ''
                Architecture = 'x64'; Confidence = 'Predicted'
            }
            $p = New-PackagerFromDrop -Analysis $hostile -PackagersRoot (Join-Path $script:genRepo 'Packagers')
            $parseErrors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($p, [ref]$null, [ref]$parseErrors)
            $parseErrors.Count | Should -Be 0
            # The hostile fragments must survive only as string content, never
            # as command names or expandable subexpressions in the AST.
            $commands = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) |
                ForEach-Object { $_.GetCommandName() }
            $commands | Should -Not -Contain 'Invoke-InjectedMarker'
            $commands | Should -Not -Contain 'Invoke-OtherMarker'
        }

        It 'a comment terminator in the publisher cannot break out of the doc header' {
            $hostile = [pscustomobject]@{
                Path = 'x'; FileName = 'evil2.msi'; InstallerType = 'MSI'
                AppName = 'CleanApp'; Publisher = 'Acme #> Invoke-InjectedMarker <#'
                SoftwareVersion = '1.0'; ProductCode = '{11111111-2222-3333-4444-555555555555}'
                InstallArgs = '/qn'; UninstallArgs = ''
                InstallCommand = ''; UninstallCommand = ''; UninstallRegistryKey = ''
                Architecture = 'x64'; Confidence = 'Authoritative'
            }
            $p = New-PackagerFromDrop -Analysis $hostile -PackagersRoot (Join-Path $script:genRepo 'Packagers')
            $parseErrors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($p, [ref]$null, [ref]$parseErrors)
            $parseErrors.Count | Should -Be 0
            $commands = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) |
                ForEach-Object { $_.GetCommandName() }
            $commands | Should -Not -Contain 'Invoke-InjectedMarker'
        }
    }
}

Describe 'Condition templates' {
    It 'ships three defaults with stable ids' {
        $doc = Get-DefaultConditionTemplates
        @($doc.Conditions).Count | Should -Be 3
        @($doc.Conditions | ForEach-Object { $_.Id }) | Should -Be @('cpu-arch', 'os-language', 'vpn-connected')
    }

    It 'default rule types map each condition onto a supported rule cmdlet' {
        $doc = Get-DefaultConditionTemplates
        foreach ($c in $doc.Conditions) {
            $c.RuleType | Should -BeIn @('CommonValue', 'OperatingSystemLanguage', 'Boolean')
        }
    }

    It 'cpu-arch maps the friendly values onto processor architecture codes' {
        $arch = (Get-DefaultConditionTemplates).Conditions | Where-Object { $_.Id -eq 'cpu-arch' }
        $arch.Values.x64   | Should -Be '9'
        $arch.Values.ARM64 | Should -Be '12'
    }

    It 'os-language pins PlatformType 1 because the stock name matches Windows and Mobile conditions' {
        $lang = (Get-DefaultConditionTemplates).Conditions | Where-Object { $_.Id -eq 'os-language' }
        $lang.PlatformType | Should -Be 1
    }

    It 'falls back to defaults when the override file is absent' {
        Mock Get-ConditionTemplatesPath -ModuleName AppPackagerCommon { Join-Path $TestDrive 'missing.json' }
        @((Get-ConditionTemplates).Conditions).Count | Should -Be 3
    }

    It 'prefers a parseable override file' {
        Mock Get-ConditionTemplatesPath -ModuleName AppPackagerCommon { Join-Path $TestDrive 'ct.json' }
        '{"SchemaVersion":1,"Conditions":[{"Id":"custom-only","Kind":"BuiltIn","RuleType":"Boolean","GlobalConditionName":"X"}]}' |
            Set-Content -Path (Join-Path $TestDrive 'ct.json')
        $doc = Get-ConditionTemplates
        @($doc.Conditions).Count | Should -Be 1
        $doc.Conditions[0].Id | Should -Be 'custom-only'
    }

    It 'falls back to defaults on a malformed override file' {
        Mock Get-ConditionTemplatesPath -ModuleName AppPackagerCommon { Join-Path $TestDrive 'bad.json' }
        Mock Write-Log -ModuleName AppPackagerCommon { }
        'not json at all {' | Set-Content -Path (Join-Path $TestDrive 'bad.json')
        @((Get-ConditionTemplates).Conditions).Count | Should -Be 3
    }

    It 'save round-trips through the override path' {
        Mock Get-ConditionTemplatesPath -ModuleName AppPackagerCommon { Join-Path $TestDrive 'rt.json' }
        $doc = Get-DefaultConditionTemplates
        $vpn = $doc.Conditions | Where-Object { $_.Id -eq 'vpn-connected' }
        $vpn.AdapterPatterns = @('CustomVpnClient')
        $null = Save-ConditionTemplates -Templates $doc
        $reloaded = Get-ConditionTemplates
        @(($reloaded.Conditions | Where-Object { $_.Id -eq 'vpn-connected' }).AdapterPatterns) | Should -Be @('CustomVpnClient')
    }
}

Describe 'New-VpnConditionScriptText' {
    It 'generates a script that parses cleanly' {
        $text = New-VpnConditionScriptText -AdapterPatterns @('Zscaler', 'Juniper') -AliasPattern 'vpn'
        $parseErrors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$null, [ref]$parseErrors)
        $parseErrors.Count | Should -Be 0
    }

    It 'embeds every pattern and escapes single quotes' {
        $text = New-VpnConditionScriptText -AdapterPatterns @('Zscaler', "O'Brien VPN") -AliasPattern 'vpn'
        $text | Should -Match "'Zscaler'"
        $text | Should -Match "'O''Brien VPN'"
    }

    It 'omits the alias check when AliasPattern is empty' {
        $text = New-VpnConditionScriptText -AdapterPatterns @('Zscaler') -AliasPattern ''
        $text | Should -Not -Match 'InterfaceAlias'
    }

    It 'executes and returns exactly one boolean' {
        $text = New-VpnConditionScriptText -AdapterPatterns @('NoSuchAdapterPatternZZZ') -AliasPattern 'nosuchaliaszzz'
        $out = @(& ([scriptblock]::Create($text)))
        $out.Count | Should -Be 1
        $out[0] | Should -BeOfType [bool]
    }
}

Describe 'Get-DeploymentTypeRequirementSpecs' {
    BeforeEach {
        $script:__savedReqEnv = $env:APP_PACKAGER_REQUIREMENTS
        Remove-Item Env:APP_PACKAGER_REQUIREMENTS -ErrorAction SilentlyContinue
    }
    AfterEach {
        if ($null -ne $script:__savedReqEnv) { $env:APP_PACKAGER_REQUIREMENTS = $script:__savedReqEnv }
        else { Remove-Item Env:APP_PACKAGER_REQUIREMENTS -ErrorAction SilentlyContinue }
    }

    It 'returns empty with no manifest field and no env var' {
        @(Get-DeploymentTypeRequirementSpecs -Manifest ([pscustomobject]@{ AppName = 'x' })).Count | Should -Be 0
    }

    It 'reads manifest Requirements' {
        $m = [pscustomobject]@{ Requirements = @([pscustomobject]@{ ConditionId = 'cpu-arch'; Value = 'x64' }) }
        $specs = @(Get-DeploymentTypeRequirementSpecs -Manifest $m)
        $specs.Count | Should -Be 1
        $specs[0].ConditionId | Should -Be 'cpu-arch'
    }

    It 'merges manifest and environment specs, manifest first' {
        $env:APP_PACKAGER_REQUIREMENTS = '{"SchemaVersion":1,"Rules":[{"ConditionId":"vpn-connected","Value":true}]}'
        $m = [pscustomobject]@{ Requirements = @([pscustomobject]@{ ConditionId = 'os-language'; Cultures = @('de-DE') }) }
        $specs = @(Get-DeploymentTypeRequirementSpecs -Manifest $m)
        @($specs | ForEach-Object { $_.ConditionId }) | Should -Be @('os-language', 'vpn-connected')
    }

    It 'preserves boolean type through the environment JSON' {
        $env:APP_PACKAGER_REQUIREMENTS = '{"SchemaVersion":1,"Rules":[{"ConditionId":"vpn-connected","Value":false}]}'
        $specs = @(Get-DeploymentTypeRequirementSpecs -Manifest ([pscustomobject]@{ AppName = 'x' }))
        $specs[0].Value | Should -BeOfType [bool]
        $specs[0].Value | Should -BeFalse
    }

    It 'throws on malformed environment JSON instead of packaging without rules' {
        $env:APP_PACKAGER_REQUIREMENTS = 'not json'
        { Get-DeploymentTypeRequirementSpecs -Manifest ([pscustomobject]@{ AppName = 'x' }) } | Should -Throw '*not valid JSON*'
    }
}

Describe 'Get-ManifestDeploymentTypeSpecs' {
    BeforeAll {
        $script:baseManifest = [pscustomobject]@{
            AppName            = 'App'
            Detection          = [pscustomobject]@{ Type = 'RegistryKeyValue'; KeyPath = 'HKLM:\X'; ValueName = 'V'; ExpectedValue = '1' }
            InstallCommandLine = ''
        }
    }

    It 'yields one spec matching current behavior when DeploymentTypes is absent' {
        $s = @(Get-ManifestDeploymentTypeSpecs -Manifest $baseManifest -NetworkContentPath '\srv\c$\app\1.0' -AppName 'App')
        $s.Count | Should -Be 1
        $s[0].DtName | Should -Be 'App'
        $s[0].ContentLocation | Should -Be '\srv\c$\app\1.0'
        $s[0].InstallCommand | Should -Be 'install.bat'
        $s[0].RequirementSource | Should -Be $baseManifest
    }

    It 'yields ordered specs with per-entry name, content, commands, and requirements' {
        $m = [pscustomobject]@{
            AppName   = 'App'
            Detection = $baseManifest.Detection
            DeploymentTypes = @(
                [pscustomobject]@{ NameSuffix = 'ARM64'; ContentSubpath = 'arm64'; Requirements = @([pscustomobject]@{ ConditionId = 'cpu-arch'; Value = 'ARM64' }) },
                [pscustomobject]@{ NameSuffix = 'x64'; InstallCommandLine = 'setup.exe /s' }
            )
        }
        $s = @(Get-ManifestDeploymentTypeSpecs -Manifest $m -NetworkContentPath '\srv\c$\app\1.0' -AppName 'App')
        $s.Count | Should -Be 2
        $s[0].DtName | Should -Be 'App - ARM64'
        $s[0].ContentLocation | Should -Be '\srv\c$\app\1.0\arm64'
        @($s[0].RequirementSource.Requirements).Count | Should -Be 1
        $s[1].DtName | Should -Be 'App - x64'
        $s[1].ContentLocation | Should -Be '\srv\c$\app\1.0'
        $s[1].InstallCommand | Should -Be 'setup.exe /s'
        @($s[1].RequirementSource.Requirements).Count | Should -Be 0
        $s[1].Detection.Type | Should -Be 'RegistryKeyValue'
    }

    It 'throws on a missing NameSuffix, duplicate names, and missing detection' {
        { Get-ManifestDeploymentTypeSpecs -Manifest ([pscustomobject]@{ AppName='App'; Detection=$baseManifest.Detection; DeploymentTypes=@([pscustomobject]@{ ContentSubpath='x' }) }) -NetworkContentPath '\s\c' -AppName 'App' } | Should -Throw '*NameSuffix*'
        { Get-ManifestDeploymentTypeSpecs -Manifest ([pscustomobject]@{ AppName='App'; Detection=$baseManifest.Detection; DeploymentTypes=@([pscustomobject]@{ NameSuffix='A' },[pscustomobject]@{ NameSuffix='a' }) }) -NetworkContentPath '\s\c' -AppName 'App' } | Should -Throw '*duplicate*'
        { Get-ManifestDeploymentTypeSpecs -Manifest ([pscustomobject]@{ AppName='App'; DeploymentTypes=@([pscustomobject]@{ NameSuffix='A' }) }) -NetworkContentPath '\s\c' -AppName 'App' } | Should -Throw '*Detection*'
    }

    It 'entry Detection overrides the base Detection' {
        $m = [pscustomobject]@{
            AppName   = 'App'
            Detection = $baseManifest.Detection
            DeploymentTypes = @([pscustomobject]@{ NameSuffix = 'S'; Detection = [pscustomobject]@{ Type = 'Script'; ScriptText = 'exit 0' } })
        }
        $s = @(Get-ManifestDeploymentTypeSpecs -Manifest $m -NetworkContentPath '\s\c' -AppName 'App')
        $s[0].Detection.Type | Should -Be 'Script'
    }
}

Describe 'Requirement spec environment exclusion' {
    It 'IgnoreEnvironment drops APP_PACKAGER_REQUIREMENTS' {
        $env:APP_PACKAGER_REQUIREMENTS = '{"SchemaVersion":1,"Rules":[{"ConditionId":"cpu-arch","Value":"x64"}]}'
        try {
            @(Get-DeploymentTypeRequirementSpecs -Manifest $null).Count | Should -Be 1
            @(Get-DeploymentTypeRequirementSpecs -Manifest $null -IgnoreEnvironment).Count | Should -Be 0
        }
        finally { $env:APP_PACKAGER_REQUIREMENTS = '' }
    }
}

Describe 'Get-RequestedPackagerVariants' {
    AfterEach { $env:APP_PACKAGER_VARIANTS = '' }

    It 'returns null when the environment variable is unset' {
        $env:APP_PACKAGER_VARIANTS = ''
        Get-RequestedPackagerVariants | Should -BeNullOrEmpty
    }

    It 'parses an Architecture split' {
        $env:APP_PACKAGER_VARIANTS = '{"SchemaVersion":1,"Split":"Architecture"}'
        $v = Get-RequestedPackagerVariants
        $v.Split | Should -Be 'Architecture'
    }

    It 'parses a Language split with cultures' {
        $env:APP_PACKAGER_VARIANTS = '{"SchemaVersion":1,"Split":"Language","Languages":["de-DE","fr-FR"]}'
        $v = Get-RequestedPackagerVariants
        $v.Split | Should -Be 'Language'
        @($v.Languages) | Should -Be @('de-DE', 'fr-FR')
    }

    It 'throws on malformed JSON, unknown split, and language split without languages' {
        $env:APP_PACKAGER_VARIANTS = '{not json'
        { Get-RequestedPackagerVariants } | Should -Throw '*not valid JSON*'
        $env:APP_PACKAGER_VARIANTS = '{"Split":"Sideways"}'
        { Get-RequestedPackagerVariants } | Should -Throw '*Split must be*'
        $env:APP_PACKAGER_VARIANTS = '{"Split":"Language"}'
        { Get-RequestedPackagerVariants } | Should -Throw '*no Languages*'
    }
}

Describe 'Get-RequestedCommandOverrides' {
    AfterEach { $env:APP_PACKAGER_COMMANDS = '' }

    It 'returns null when unset' {
        $env:APP_PACKAGER_COMMANDS = ''
        Get-RequestedCommandOverrides | Should -BeNullOrEmpty
    }

    It 'parses install and uninstall overrides' {
        $env:APP_PACKAGER_COMMANDS = '{"SchemaVersion":1,"Install":"setup.exe /s","Uninstall":"setup.exe /x"}'
        $c = Get-RequestedCommandOverrides
        $c.Install | Should -Be 'setup.exe /s'
        $c.Uninstall | Should -Be 'setup.exe /x'
    }

    It 'accepts a single-sided override' {
        $env:APP_PACKAGER_COMMANDS = '{"Install":"setup.exe /s"}'
        (Get-RequestedCommandOverrides).Uninstall | Should -Be ''
    }

    It 'throws on malformed JSON and on an empty override' {
        $env:APP_PACKAGER_COMMANDS = '{nope'
        { Get-RequestedCommandOverrides } | Should -Throw '*not valid JSON*'
        $env:APP_PACKAGER_COMMANDS = '{"SchemaVersion":1}'
        { Get-RequestedCommandOverrides } | Should -Throw '*neither*'
    }
}

Describe 'Stage manifest override recording' {
    AfterEach { $env:APP_PACKAGER_COMMANDS = '' }

    It 'records active command overrides in the manifest' {
        $env:APP_PACKAGER_COMMANDS = '{"SchemaVersion":1,"Install":"setup.exe /s"}'
        $dir = Join-Path $TestDrive 'ov-stage'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Set-Content (Join-Path $dir 'payload.txt') 'x'
        $mp = Join-Path $dir 'stage-manifest.json'
        Write-StageManifest -Path $mp -ManifestData @{ AppName = 'A'; Detection = @{ Type = 'Script'; ScriptText = 'exit 0' } }
        $doc = Get-Content $mp -Raw | ConvertFrom-Json
        $doc.CommandOverrides.Install | Should -Be 'setup.exe /s'
    }

    It 'writes no override field on a stock build' {
        $env:APP_PACKAGER_COMMANDS = ''
        $dir = Join-Path $TestDrive 'stock-stage'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Set-Content (Join-Path $dir 'payload.txt') 'x'
        $mp = Join-Path $dir 'stage-manifest.json'
        Write-StageManifest -Path $mp -ManifestData @{ AppName = 'A'; Detection = @{ Type = 'Script'; ScriptText = 'exit 0' } }
        $doc = Get-Content $mp -Raw | ConvertFrom-Json
        $doc.PSObject.Properties['CommandOverrides'] | Should -BeNullOrEmpty
    }
}

Describe 'Intune Win32 publishing' {
    BeforeAll {
        # Minimal but structurally valid .intunewin: outer zip with
        # Metadata/Detection.xml + Contents/IntunePackage.intunewin.
        Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
        $script:iwPath = Join-Path $TestDrive 'fixture.intunewin'
        $payloadBytes = [Text.Encoding]::ASCII.GetBytes('ENCRYPTED-PAYLOAD-BYTES-0123456789')
        $detectionXml = @'
<ApplicationInfo xmlns:xsd="http://www.w3.org/2001/XMLSchema" ToolVersion="1.8.6.0">
  <Name>Fixture App</Name>
  <UnencryptedContentSize>12345</UnencryptedContentSize>
  <FileName>IntunePackage.intunewin</FileName>
  <SetupFile>install.bat</SetupFile>
  <EncryptionInfo>
    <EncryptionKey>a2V5a2V5a2V5a2V5a2V5a2V5a2V5a2V5a2V5a2V5a2U9</EncryptionKey>
    <MacKey>bWFja2V5bWFja2V5bWFja2V5bWFja2V5bWFja2V5bWE9</MacKey>
    <InitializationVector>aXZpdml2aXZpdml2aXZpdg==</InitializationVector>
    <Mac>bWFjbWFjbWFjbWFjbWFjbWFjbWFjbWFjbWFjbWFjbWE9</Mac>
    <ProfileIdentifier>ProfileVersion1</ProfileIdentifier>
    <FileDigest>ZGlnZXN0ZGlnZXN0ZGlnZXN0ZGlnZXN0ZGlnZXN0ZGk9</FileDigest>
    <FileDigestAlgorithm>SHA256</FileDigestAlgorithm>
  </EncryptionInfo>
</ApplicationInfo>
'@
        $fs = [System.IO.File]::Create($script:iwPath)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        $e1 = $zip.CreateEntry('IntuneWinPackage/Metadata/Detection.xml')
        $w = New-Object System.IO.StreamWriter($e1.Open()); $w.Write($detectionXml); $w.Dispose()
        $e2 = $zip.CreateEntry('IntuneWinPackage/Contents/IntunePackage.intunewin')
        $s = $e2.Open(); $s.Write($payloadBytes, 0, $payloadBytes.Length); $s.Dispose()
        $zip.Dispose(); $fs.Dispose()
    }

    It 'reads metadata and encryption info from the package' {
        $m = Get-IntuneWinEncryptionInfo -Path $script:iwPath
        $m.SetupFile | Should -Be 'install.bat'
        $m.UnencryptedContentSize | Should -Be 12345
        $m.ProfileIdentifier | Should -Be 'ProfileVersion1'
        $m.FileDigestAlgorithm | Should -Be 'SHA256'
        $m.EncryptionKey | Should -Not -BeNullOrEmpty
    }

    It 'extracts the encrypted payload' {
        $out = Join-Path $TestDrive 'payload.bin'
        $p = Export-IntuneWinPayload -Path $script:iwPath -Destination $out
        $p.Size | Should -Be 34
        [Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($out)) | Should -Be 'ENCRYPTED-PAYLOAD-BYTES-0123456789'
    }

    It 'throws on a file that is not an intunewin package' {
        $bad = Join-Path $TestDrive 'bad.intunewin'
        Set-Content $bad 'not a zip'
        { Get-IntuneWinEncryptionInfo -Path $bad } | Should -Throw
    }

    It 'maps RegistryKeyValue detection to a version-compare registry rule' {
        $m = [pscustomobject]@{ Detection = [pscustomobject]@{ Type='RegistryKeyValue'; RegistryKeyRelative='SOFTWARE\X'; ValueName='DisplayVersion'; ExpectedValue='26.02.00.0'; Is64Bit=$true } }
        $r = @(ConvertTo-IntuneWin32Rules -Manifest $m)
        $r.Count | Should -Be 1
        $r[0]['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppRegistryRule'
        $r[0].keyPath | Should -Be 'HKEY_LOCAL_MACHINE\SOFTWARE\X'
        $r[0].operationType | Should -Be 'version'
        $r[0].check32BitOn64System | Should -BeFalse
    }

    It 'maps non-version registry values to string compare and File version detection to a file rule' {
        $m1 = [pscustomobject]@{ Detection = [pscustomobject]@{ Type='RegistryKeyValue'; RegistryKeyRelative='SOFTWARE\X'; ValueName='Channel'; ExpectedValue='Stable'; Is64Bit=$true } }
        (@(ConvertTo-IntuneWin32Rules -Manifest $m1))[0].operationType | Should -Be 'string'
        $m2 = [pscustomobject]@{ Detection = [pscustomobject]@{ Type='File'; FilePath='C:\Program Files\App'; FileName='app.exe'; PropertyType='Version'; Operator='GreaterEquals'; ExpectedValue='1.2.3'; Is64Bit=$true } }
        $r2 = (@(ConvertTo-IntuneWin32Rules -Manifest $m2))[0]
        $r2['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppFileSystemRule'
        $r2.operator | Should -Be 'greaterThanOrEqual'
    }

    It 'maps Script detection with base64 content and refuses OR compounds' {
        $m = [pscustomobject]@{ Detection = [pscustomobject]@{ Type='Script'; ScriptText='exit 0' } }
        $r = (@(ConvertTo-IntuneWin32Rules -Manifest $m))[0]
        [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($r.scriptContent)) | Should -Be 'exit 0'
        $mOr = [pscustomobject]@{ Detection = [pscustomobject]@{ Type='Compound'; Connector='Or'; Clauses=@() } }
        { ConvertTo-IntuneWin32Rules -Manifest $mOr } | Should -Throw '*OR-connected*'
    }

    It 'routes an HKCU manifest hive into the registry rule keyPath' {
        $m = [pscustomobject]@{ Detection = [pscustomobject]@{ Type='RegistryKeyValue'; Hive='CurrentUser'; RegistryKeyRelative='Software\Microsoft\Windows\CurrentVersion\Uninstall\X'; ValueName='DisplayVersion'; ExpectedValue='1.0.0'; Is64Bit=$true } }
        (@(ConvertTo-IntuneWin32Rules -Manifest $m))[0].keyPath | Should -Be 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\X'
        $m2 = [pscustomobject]@{ Detection = [pscustomobject]@{ Type='RegistryKey'; RegistryKeyRelative='SOFTWARE\Y'; Is64Bit=$true } }
        (@(ConvertTo-IntuneWin32Rules -Manifest $m2))[0].keyPath | Should -Be 'HKEY_LOCAL_MACHINE\SOFTWARE\Y'
    }

    It 'maps AND compounds clause-per-rule' {
        $m = [pscustomobject]@{ Detection = [pscustomobject]@{ Type='Compound'; Connector='And'; Clauses=@(
            [pscustomobject]@{ Type='RegistryKey'; RegistryKeyRelative='SOFTWARE\A'; Is64Bit=$true },
            [pscustomobject]@{ Type='File'; FilePath='C:\B'; FileName='b.dll'; PropertyType='Existence'; Is64Bit=$true }
        ) } }
        @(ConvertTo-IntuneWin32Rules -Manifest $m).Count | Should -Be 2
    }

    It 'runs the full publish flow against mocked Graph calls' {
        $global:graphCalls = [System.Collections.Generic.List[object]]::new()
        Mock Get-MsGraphToken { 'tok' } -ModuleName AppPackagerCommon
        Mock Invoke-AzureBlobUpload { } -ModuleName AppPackagerCommon
        Mock Invoke-GraphJson {
            $call = @{ Method = $Method; Uri = $Uri; Body = $Body }
            ([System.Collections.Generic.List[object]]$global:graphCalls).Add($call)
            $key = "$Method $Uri"
            if ($key -match 'GET .*mobileApps\?')        { return [pscustomobject]@{ value = @() } }
            if ($key -match 'POST .*mobileApps$')        { return [pscustomobject]@{ id = 'app-1' } }
            if ($key -match 'POST .*contentVersions$')   { return [pscustomobject]@{ id = '1' } }
            if ($key -match 'POST .*files$')             { return [pscustomobject]@{ id = 'file-1' } }
            if ($key -match 'GET .*files/file-1$')       {
                $global:fileGets++
                if ($global:fileGets -eq 1) { return [pscustomobject]@{ uploadState = 'azureStorageUriRequestSuccess'; azureStorageUri = 'https://blob/x?sas=1' } }
                return [pscustomobject]@{ uploadState = 'commitFileSuccess'; azureStorageUri = 'https://blob/x?sas=1' }
            }
            if ($key -match 'POST .*commit$')            { return $null }
            if ($key -match 'PATCH .*mobileApps/app-1$') { return $null }
            throw "Unexpected Graph call: $key"
        } -ModuleName AppPackagerCommon

        $global:fileGets = 0
        $manifest = [pscustomobject]@{
            AppName = 'Fixture App'; Publisher = 'Fixture'
            InstallCommandLine = ''; UninstallCommandLine = ''
            Detection = [pscustomobject]@{ Type='RegistryKeyValue'; RegistryKeyRelative='SOFTWARE\X'; ValueName='DisplayVersion'; ExpectedValue='1.0'; Is64Bit=$true }
        }
        $appId = Publish-IntuneWin32App -TenantId 't' -ClientId 'c' -ClientSecret 's' -IntuneWinPath $script:iwPath -Manifest $manifest
        $appId | Should -Be 'app-1'
        Should -Invoke Invoke-AzureBlobUpload -ModuleName AppPackagerCommon -Times 1 -Exactly
        $commit = $global:graphCalls | Where-Object { $_.Uri -like '*commit' }
        $commit.Body.fileEncryptionInfo.profileIdentifier | Should -Be 'ProfileVersion1'
        $patch = $global:graphCalls | Where-Object { $_.Method -eq 'PATCH' }
        $patch.Body.committedContentVersion | Should -Be '1'
    }
}

Describe 'Intune Win32 repeat publish' {
    BeforeAll {
        Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
        $script:iw2Path = Join-Path $TestDrive 'fixture2.intunewin'
        $xml = '<ApplicationInfo><Name>F</Name><UnencryptedContentSize>10</UnencryptedContentSize><FileName>IntunePackage.intunewin</FileName><SetupFile>install.bat</SetupFile><EncryptionInfo><EncryptionKey>k</EncryptionKey><MacKey>m</MacKey><InitializationVector>i</InitializationVector><Mac>c</Mac><ProfileIdentifier>ProfileVersion1</ProfileIdentifier><FileDigest>d</FileDigest><FileDigestAlgorithm>SHA256</FileDigestAlgorithm></EncryptionInfo></ApplicationInfo>'
        $fs = [System.IO.File]::Create($script:iw2Path)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        $e1 = $zip.CreateEntry('IntuneWinPackage/Metadata/Detection.xml')
        $w = New-Object System.IO.StreamWriter($e1.Open()); $w.Write($xml); $w.Dispose()
        $e2 = $zip.CreateEntry('IntuneWinPackage/Contents/IntunePackage.intunewin')
        $s = $e2.Open(); $b = [Text.Encoding]::ASCII.GetBytes('X'); $s.Write($b, 0, 1); $s.Dispose()
        $zip.Dispose(); $fs.Dispose()
    }

    It 'updates the existing app instead of creating a duplicate, and escapes quotes in the filter' {
        $global:g2 = [System.Collections.Generic.List[object]]::new()
        $global:g2Gets = 0
        Mock Get-MsGraphToken { 'tok' } -ModuleName AppPackagerCommon
        Mock Invoke-AzureBlobUpload { } -ModuleName AppPackagerCommon
        Mock Invoke-GraphJson {
            ([System.Collections.Generic.List[object]]$global:g2).Add(@{ Method = $Method; Uri = $Uri; Body = $Body })
            $key = "$Method $Uri"
            if ($key -match "GET .*mobileApps\?") { return [pscustomobject]@{ value = @([pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'existing-1' }) } }
            if ($key -match 'PATCH .*mobileApps/existing-1$') { return $null }
            if ($key -match 'POST .*contentVersions$')   { return [pscustomobject]@{ id = '7' } }
            if ($key -match 'POST .*files$')             { return [pscustomobject]@{ id = 'f' } }
            if ($key -match 'GET .*files/f$')            {
                $global:g2Gets++
                if ($global:g2Gets -eq 1) { return [pscustomobject]@{ uploadState = 'azureStorageUriRequestSuccess'; azureStorageUri = 'https://blob/y?sas=1' } }
                return [pscustomobject]@{ uploadState = 'commitFileSuccess'; azureStorageUri = 'https://blob/y?sas=1' }
            }
            if ($key -match 'POST .*commit$')            { return $null }
            throw "Unexpected Graph call: $key"
        } -ModuleName AppPackagerCommon

        $manifest = [pscustomobject]@{
            AppName = "O'Reilly App"; Publisher = 'P'
            InstallCommandLine = ''; UninstallCommandLine = ''
            Detection = [pscustomobject]@{ Type='Script'; ScriptText='exit 0' }
        }
        $appId = Publish-IntuneWin32App -TenantId 't' -ClientId 'c' -ClientSecret 's' -IntuneWinPath $script:iw2Path -Manifest $manifest
        $appId | Should -Be 'existing-1'
        # the display-name fallback filter escaped the apostrophe
        $filterCall = @($global:g2 | Where-Object { $_.Method -eq 'GET' -and $_.Uri -like '*displayName eq*' })[0]
        $filterCall.Uri | Should -BeLike "*O''Reilly App*"
        # no create POST to the mobileApps collection happened
        @($global:g2 | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match 'mobileApps$' }).Count | Should -Be 0
        # two PATCHes: metadata update + committedContentVersion
        @($global:g2 | Where-Object { $_.Method -eq 'PATCH' }).Count | Should -Be 2
    }
}


Describe 'Get-GitHubApiCurlArgs' {
    BeforeEach {
        $script:savedGitHubToken = $env:GITHUB_TOKEN
        $script:savedGhToken = $env:GH_TOKEN
        $env:GITHUB_TOKEN = $null
        $env:GH_TOKEN = $null
        # Forget any CLI token cached by an earlier call in this process.
        InModuleScope AppPackagerCommon { $script:GitHubCliToken = $null }
    }
    AfterEach {
        $env:GITHUB_TOKEN = $script:savedGitHubToken
        $env:GH_TOKEN = $script:savedGhToken
        InModuleScope AppPackagerCommon { $script:GitHubCliToken = $null }
    }

    It 'sends GITHUB_TOKEN as a bearer token' {
        $env:GITHUB_TOKEN = 'token-a'
        $args = Get-GitHubApiCurlArgs
        $args | Should -Be @('-H', 'Authorization: Bearer token-a', '-H', 'X-GitHub-Api-Version: 2022-11-28')
    }

    It 'falls back to GH_TOKEN' {
        $env:GH_TOKEN = 'token-b'
        (Get-GitHubApiCurlArgs)[1] | Should -Be 'Authorization: Bearer token-b'
    }

    It 'returns no arguments without a token or a signed-in GitHub CLI' {
        Mock Get-Command { $null } -ModuleName AppPackagerCommon -ParameterFilter { $Name -eq 'gh.exe' }
        @(Get-GitHubApiCurlArgs).Count | Should -Be 0
    }

    It 'uses the GitHub CLI login when no environment token is set' {
        Mock Get-Command { [pscustomobject]@{ Source = 'gh.exe' } } -ModuleName AppPackagerCommon -ParameterFilter { $Name -eq 'gh.exe' }
        InModuleScope AppPackagerCommon { $script:GitHubCliToken = 'cli-token' }
        (Get-GitHubApiCurlArgs)[1] | Should -Be 'Authorization: Bearer cli-token'
    }
}

Describe 'Get-RequestedInstallMode' {
    AfterEach { Remove-Item Env:\APP_PACKAGER_INSTALL_MODE -ErrorAction SilentlyContinue }

    It 'returns $null when nothing is requested' {
        Remove-Item Env:\APP_PACKAGER_INSTALL_MODE -ErrorAction SilentlyContinue
        Get-RequestedInstallMode | Should -BeNullOrEmpty
    }

    It 'returns the requested mode' {
        $env:APP_PACKAGER_INSTALL_MODE = 'CurrentUser'
        Get-RequestedInstallMode | Should -Be 'CurrentUser'
    }

    It 'throws on an unknown value' {
        $env:APP_PACKAGER_INSTALL_MODE = 'Everyone'
        { Get-RequestedInstallMode } | Should -Throw
    }
}

Describe 'Set-StageManifestInstallMode' {
    BeforeAll {
        $script:modeInstaller = New-InnoFixture -Path (Join-Path $TestDrive 'ModeTree-Setup.exe') -AppId 'ModeTree' -DefaultDirName '{autopf}\Mode Tree' -Privileges 3 -Overrides 1
        $script:fixedInstaller = New-InnoFixture -Path (Join-Path $TestDrive 'ModeFixed-Setup.exe') -AppId 'ModeFixed' -DefaultDirName '{autopf}\Mode Fixed' -Privileges 2

        # A stage folder the way a packager with standard wrappers leaves it.
        function script:New-ModeStage {
            param([string]$Name, [string]$Installer, [string]$InstallArgs, [string]$UninstallArgs, [hashtable]$Detection)
            $root = Join-Path $TestDrive $Name
            New-Item -ItemType Directory -Path $root -Force | Out-Null
            Copy-Item -LiteralPath $Installer -Destination (Join-Path $root (Split-Path -Leaf $Installer)) -Force
            $w = New-ExeWrapperContent -InstallerFileName (Split-Path -Leaf $Installer) -InstallArgs ("'" + $InstallArgs + "'") -UninstallCommand 'C:\Program Files\Mode Tree\unins000.exe' -UninstallArgs ("'" + $UninstallArgs + "'")
            Write-ContentWrappers -OutputPath $root -InstallPs1Content $w.Install -UninstallPs1Content $w.Uninstall
            $manifest = @{
                AppName = 'Mode Tree'; Publisher = 'Test'; SoftwareVersion = '1.0'
                InstallerFile = (Split-Path -Leaf $Installer); InstallerType = 'EXE'
                InstallArgs = $InstallArgs; UninstallArgs = $UninstallArgs
                UninstallCommand = 'C:\Program Files\Mode Tree\unins000.exe'
                Detection = $Detection
            }
            return @{ Root = $root; Manifest = $manifest }
        }
    }

    It 'moves an all-users package to the per-user branch: switch, uninstaller, hive and behavior' {
        $s = New-ModeStage -Name 'mode-user' -Installer $script:modeInstaller -InstallArgs '/VERYSILENT /NORESTART /ALLUSERS /SP-' -UninstallArgs '/VERYSILENT /NORESTART' `
            -Detection @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ModeTree_is1'; ValueName = 'DisplayVersion'; ExpectedValue = '1.0'; Is64Bit = $true }
        $applied = Set-StageManifestInstallMode -StageRoot $s.Root -ManifestData $s.Manifest -Mode CurrentUser
        $applied | Should -BeTrue
        $m = $s.Manifest
        $m.InstallArgs | Should -Be '/VERYSILENT /NORESTART /SP- /CURRENTUSER'
        $m.UninstallArgs | Should -Be '/VERYSILENT /NORESTART'
        $m.UninstallCommand | Should -Be '%LOCALAPPDATA%\Programs\Mode Tree\unins000.exe'
        $m.Detection.Hive | Should -Be 'CurrentUser'
        $m.Detection.RegistryKeyRelative | Should -Be 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ModeTree_is1'
        $m.InstallationBehaviorType | Should -Be 'InstallForUser'
        $m.LogonRequirementType | Should -Be 'OnlyWhenUserLoggedOn'
        $m.InstallMode | Should -Be 'CurrentUser'
        $m.InstallContext | Should -Be 'PerUser'
        (Get-Content (Join-Path $s.Root 'install.ps1') -Raw) | Should -Match ([regex]::Escape("@('/VERYSILENT /NORESTART /SP- /CURRENTUSER')"))
        (Get-Content (Join-Path $s.Root 'uninstall.ps1') -Raw) | Should -Match ([regex]::Escape('%LOCALAPPDATA%\Programs\Mode Tree\unins000.exe'))
    }

    It 'moves a per-user package to the all-users branch and maps a file detection path' {
        $s = New-ModeStage -Name 'mode-system' -Installer $script:modeInstaller -InstallArgs '/VERYSILENT /SP- /CURRENTUSER' -UninstallArgs '/VERYSILENT' `
            -Detection @{ Type = 'File'; FilePath = '%LOCALAPPDATA%\Programs\Mode Tree'; FileName = 'tree.exe'; PropertyType = 'Version'; ExpectedValue = '1.0'; Is64Bit = $true }
        [void](Set-StageManifestInstallMode -StageRoot $s.Root -ManifestData $s.Manifest -Mode AllUsers)
        $m = $s.Manifest
        $m.InstallArgs | Should -Be '/VERYSILENT /SP- /ALLUSERS'
        $m.UninstallCommand | Should -Be '%ProgramFiles%\Mode Tree\unins000.exe'
        $m.Detection.FilePath | Should -Be '%ProgramFiles%\Mode Tree'
        $m.InstallationBehaviorType | Should -Be 'InstallForSystem'
        $m.ContainsKey('LogonRequirementType') | Should -BeFalse
        $m.InstallContext | Should -Be 'PerMachine'
    }

    It 'converts file detection to the analyzed HKCU uninstall version for user mode' {
        $s = New-ModeStage -Name 'mode-user-file' -Installer $script:modeInstaller -InstallArgs '/VERYSILENT /ALLUSERS' -UninstallArgs '/VERYSILENT' `
            -Detection @{ Type = 'File'; FilePath = '%ProgramFiles%\Mode Tree'; FileName = 'tree.exe'; PropertyType = 'Version'; ExpectedValue = '9.9'; Is64Bit = $true }
        [void](Set-StageManifestInstallMode -StageRoot $s.Root -ManifestData $s.Manifest -Mode CurrentUser)
        $s.Manifest.Detection.Type | Should -Be 'RegistryKeyValue'
        $s.Manifest.Detection.Hive | Should -Be 'CurrentUser'
        $s.Manifest.Detection.RegistryKeyRelative | Should -Be 'Software\Microsoft\Windows\CurrentVersion\Uninstall\ModeTree_is1'
        $s.Manifest.Detection.ValueName | Should -Be 'DisplayVersion'
        $s.Manifest.Detection.PropertyType | Should -Be 'Version'
        $s.Manifest.Detection.Operator | Should -Be 'GreaterEquals'
        $s.Manifest.Detection.ExpectedValue | Should -Be '1.0'
        $s.Manifest.Detection.ContainsKey('FilePath') | Should -BeFalse
    }

    It 'rejects user file detection without a concrete HKCU uninstall key' {
        $s = New-ModeStage -Name 'mode-user-unknown-key' -Installer $script:modeInstaller -InstallArgs '/VERYSILENT' -UninstallArgs '/VERYSILENT' `
            -Detection @{ Type = 'File'; FilePath = '%ProgramFiles%\Mode Tree'; FileName = 'tree.exe'; PropertyType = 'Version'; ExpectedValue = '1.0' }
        Mock Set-InstallerAnalysisMode -ModuleName AppPackagerCommon {
            [pscustomobject]@{ InstallContext = 'PerUser'; UninstallRegistryHive = 'HKCU'; UninstallRegistryKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_ID}'; UninstallCommand = '' }
        }
        { Set-StageManifestInstallMode -StageRoot $s.Root -ManifestData $s.Manifest -Mode CurrentUser } | Should -Throw '*concrete HKCU uninstall key*'
    }

    It 'allows literal GUID uninstall keys for user file detection' {
        $s = New-ModeStage -Name 'mode-user-guid-key' -Installer $script:modeInstaller -InstallArgs '/VERYSILENT' -UninstallArgs '/VERYSILENT' `
            -Detection @{ Type = 'File'; FilePath = '%ProgramFiles%\Mode Tree'; FileName = 'tree.exe'; PropertyType = 'Version'; ExpectedValue = '1.0' }
        Mock Set-InstallerAnalysisMode -ModuleName AppPackagerCommon {
            [pscustomobject]@{ InstallContext = 'PerUser'; UninstallRegistryHive = 'HKCU'; UninstallRegistryKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{771FD6B0-FA20-440A-A002-3B3BAC16DC50}_is1'; UninstallCommand = '' }
        }
        [void](Set-StageManifestInstallMode -StageRoot $s.Root -ManifestData $s.Manifest -Mode CurrentUser)
        $s.Manifest.Detection.RegistryKeyRelative | Should -Be 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{771FD6B0-FA20-440A-A002-3B3BAC16DC50}_is1'
    }

    It 'throws when the installer offers no mode switch' {
        $s = New-ModeStage -Name 'mode-fixed' -Installer $script:fixedInstaller -InstallArgs '/VERYSILENT' -UninstallArgs '/VERYSILENT' `
            -Detection @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ModeFixed_is1'; ValueName = 'DisplayVersion'; ExpectedValue = '1.0'; Is64Bit = $true }
        { Set-StageManifestInstallMode -StageRoot $s.Root -ManifestData $s.Manifest -Mode CurrentUser } | Should -Throw '*offers*'
    }

    It 'is applied by Write-StageManifest when the environment requests a mode' {
        $s = New-ModeStage -Name 'mode-write' -Installer $script:modeInstaller -InstallArgs '/VERYSILENT /ALLUSERS' -UninstallArgs '/VERYSILENT' `
            -Detection @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ModeTree_is1'; ValueName = 'DisplayVersion'; ExpectedValue = '1.0'; Is64Bit = $true }
        $env:APP_PACKAGER_INSTALL_MODE = 'CurrentUser'
        try {
            Write-StageManifest -Path (Join-Path $s.Root 'stage-manifest.json') -ManifestData $s.Manifest -PackagerScriptPath 'C:\nowhere\package-modetree.ps1'
        }
        finally { Remove-Item Env:\APP_PACKAGER_INSTALL_MODE -ErrorAction SilentlyContinue }
        $read = Read-StageManifest -Path (Join-Path $s.Root 'stage-manifest.json')
        $read.InstallMode | Should -Be 'CurrentUser'
        $read.InstallationBehaviorType | Should -Be 'InstallForUser'
        $read.Detection.Hive | Should -Be 'CurrentUser'
        @($read.FileHashes | Where-Object { $_.RelativePath -eq 'install.ps1' }).Count | Should -Be 1
    }
}

# ---------------------------------------------------------------------------
# Intune rule mapping
# ---------------------------------------------------------------------------

Describe 'ConvertTo-IntuneWin32Rules operators' {
    BeforeAll {
        function New-RuleManifest { param($Detection) [pscustomobject]@{ Detection = [pscustomobject]$Detection } }
    }

    It 'maps GreaterEquals on a version-shaped value to a version rule' {
        $r = @(ConvertTo-IntuneWin32Rules -Manifest (New-RuleManifest @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'K'; ExpectedValue = '1.2.3'; Operator = 'GreaterEquals'; Is64Bit = $true }))
        $r[0].operator | Should -Be 'greaterThanOrEqual'
        $r[0].operationType | Should -Be 'version'
    }

    It 'defaults a missing registry operator to equality and a missing file operator to GreaterEquals' {
        $reg = @(ConvertTo-IntuneWin32Rules -Manifest (New-RuleManifest @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'K'; ExpectedValue = 'abc'; Is64Bit = $true }))
        $reg[0].operator | Should -Be 'equal'
        $file = @(ConvertTo-IntuneWin32Rules -Manifest (New-RuleManifest @{ Type = 'File'; FilePath = 'C:\x'; FileName = 'a.exe'; PropertyType = 'Version'; ExpectedValue = '1.0'; Is64Bit = $true }))
        $file[0].operator | Should -Be 'greaterThanOrEqual'
    }

    It 'refuses to convert a prefix comparison to a script without the profile opt-in' {
        { ConvertTo-IntuneWin32Rules -Manifest (New-RuleManifest @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'K'; ValueName = 'ProductName'; Operator = 'BeginsWith'; ExpectedValue = 'Windows'; Is64Bit = $true }) } |
            Should -Throw '*IntuneScriptConversion*'
    }

    It 'turns BeginsWith into a script rule that keeps the prefix comparison' {
        $r = @(ConvertTo-IntuneWin32Rules -Manifest (New-RuleManifest @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'; ValueName = 'ProductName'; Operator = 'BeginsWith'; ExpectedValue = 'Windows'; Is64Bit = $true; IntuneScriptConversion = $true }))
        $r.Count | Should -Be 1
        $r[0]['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppPowerShellScriptRule'
        $text = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($r[0].scriptContent))
        $text | Should -Match 'Registry64'
        $text | Should -Match "StartsWith\('Windows'"
        $ps = [powershell]::Create().AddScript($text)
        try { @($ps.Invoke())[0] | Should -Be 'Detected' } finally { $ps.Dispose() }
    }

    It 'escapes apostrophes in the script rule' {
        $r = @(ConvertTo-IntuneWin32Rules -Manifest (New-RuleManifest @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = "SOFTWARE\O'Vendor"; ValueName = 'DisplayVersion'; Operator = 'BeginsWith'; ExpectedValue = "17.6"; Is64Bit = $false; IntuneScriptConversion = $true }))
        $text = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($r[0].scriptContent))
        $text.Contains("OpenSubKey('SOFTWARE\O''Vendor')") | Should -BeTrue
        $text | Should -Match 'Registry32'
        $errs = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$null, [ref]$errs)
        $errs.Count | Should -Be 0
    }

    It 'refuses an operator Graph cannot express' {
        { ConvertTo-IntuneWin32Rules -Manifest (New-RuleManifest @{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'K'; ExpectedValue = '1'; Operator = 'OneOf'; Is64Bit = $true }) } |
            Should -Throw '*no Intune rule mapping*'
    }

    It 'refuses a compound that mixes a script rule with other clauses' {
        $det = @{ Type = 'Compound'; Connector = 'And'; IntuneScriptConversion = $true; Clauses = @(
            [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'K'; ExpectedValue = '1'; Operator = 'BeginsWith'; Is64Bit = $true },
            [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'K2'; Is64Bit = $true }) }
        { ConvertTo-IntuneWin32Rules -Manifest (New-RuleManifest $det) } | Should -Throw '*script rule*'
    }
}

# ---------------------------------------------------------------------------
# Azure block-blob upload
# ---------------------------------------------------------------------------

Describe 'Invoke-AzureBlobUpload' {
    It 'sends the file bytes unchanged, including a partial final chunk' {
        $payload = Join-Path $TestDrive 'payload.bin'
        $bytes = New-Object byte[] 3000
        (New-Object System.Random 7).NextBytes($bytes)
        [IO.File]::WriteAllBytes($payload, $bytes)

        $probe = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
        $probe.Start(); $port = $probe.LocalEndpoint.Port; $probe.Stop()
        $listener = New-Object Net.HttpListener
        $listener.Prefixes.Add("http://127.0.0.1:$port/")
        $listener.Start()
        $client = [powershell]::Create()
        try {
            [void]$client.AddScript({
                param($ModulePath, $PayloadPath, $UploadUri)
                Import-Module $ModulePath -Force -DisableNameChecking
                Invoke-AzureBlobUpload -Uri $UploadUri -FilePath $PayloadPath -ChunkSizeMB 1
            }).AddArgument("$PSScriptRoot\AppPackagerCommon.psd1").AddArgument($payload).AddArgument("http://127.0.0.1:$port/blob?sv=1")
            $pending = $client.BeginInvoke()
            $received = New-Object IO.MemoryStream
            $blockRequests = 0
            for ($i = 0; $i -lt 2; $i++) {
                $ctxAsync = $listener.BeginGetContext($null, $null)
                if (-not $ctxAsync.AsyncWaitHandle.WaitOne(15000)) { throw 'upload did not reach the receiver' }
                $ctx = $listener.EndGetContext($ctxAsync)
                if ($ctx.Request.QueryString['comp'] -eq 'block') { $ctx.Request.InputStream.CopyTo($received); $blockRequests++ }
                $ctx.Response.StatusCode = 201; $ctx.Response.ContentLength64 = 0; $ctx.Response.Close()
            }
            [void]$client.EndInvoke($pending)
            $client.HadErrors | Should -BeFalse
            $blockRequests | Should -Be 1
            [Convert]::ToBase64String($received.ToArray()) | Should -Be ([Convert]::ToBase64String($bytes))
        }
        finally { $client.Stop(); $client.Dispose(); $listener.Close() }
    }
}

# ---------------------------------------------------------------------------
# Conditional cache refresh
# ---------------------------------------------------------------------------

Describe 'Invoke-CachedDownload' {
    BeforeEach { $script:downloadCalls = @() }

    It 'downloads with the remote timestamp when nothing is cached' {
        Mock Invoke-DownloadWithRetry -ModuleName AppPackagerCommon { $script:downloadCalls += ,@($OutFile, $ExtraCurlArgs); Set-Content -LiteralPath $OutFile -Value 'new' }
        $target = Join-Path $TestDrive 'first.bin'
        Invoke-CachedDownload -Url 'https://example.invalid/x' -OutFile $target -Quiet
        $script:downloadCalls.Count | Should -Be 1
        $script:downloadCalls[0][0] | Should -Be $target
        $script:downloadCalls[0][1] | Should -Be @('-R')
        Get-Content -LiteralPath $target | Should -Be 'new'
    }

    It 'sends the cached timestamp and keeps the cache when the server has nothing newer' {
        Mock Invoke-DownloadWithRetry -ModuleName AppPackagerCommon { $script:downloadCalls += ,@($OutFile, $ExtraCurlArgs) }
        $target = Join-Path $TestDrive 'same.bin'
        Set-Content -LiteralPath $target -Value 'old'
        Invoke-CachedDownload -Url 'https://example.invalid/x' -OutFile $target -Quiet
        $script:downloadCalls[0][0] | Should -Be ($target + '.refresh')
        $script:downloadCalls[0][1] | Should -Be @('-R', '-z', $target)
        Get-Content -LiteralPath $target | Should -Be 'old'
        Test-Path -LiteralPath ($target + '.refresh') | Should -BeFalse
    }

    It 'replaces the cache with a completed refresh' {
        Mock Invoke-DownloadWithRetry -ModuleName AppPackagerCommon { Set-Content -LiteralPath $OutFile -Value 'newer' }
        $target = Join-Path $TestDrive 'replace.bin'
        Set-Content -LiteralPath $target -Value 'old'
        Invoke-CachedDownload -Url 'https://example.invalid/x' -OutFile $target -Quiet
        Get-Content -LiteralPath $target | Should -Be 'newer'
        Test-Path -LiteralPath ($target + '.refresh') | Should -BeFalse
    }

    It 'keeps the cache when the refresh fails' {
        Mock Invoke-DownloadWithRetry -ModuleName AppPackagerCommon { throw 'offline' }
        $target = Join-Path $TestDrive 'keep.bin'
        Set-Content -LiteralPath $target -Value 'old'
        { Invoke-CachedDownload -Url 'https://example.invalid/x' -OutFile $target -Quiet } | Should -Not -Throw
        Get-Content -LiteralPath $target | Should -Be 'old'
    }
}

# ---------------------------------------------------------------------------
# Workbench adapters: finalization hook, plan digest, schema 4
# ---------------------------------------------------------------------------

Describe 'Write-StageManifest finalization hook' {
    It 'calls the finalizer after the icon and before hashing' {
        $root = Join-Path $TestDrive 'hook-order'
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        Mock Invoke-StageFinalization -ModuleName AppPackagerCommon {
            Set-Content -LiteralPath (Join-Path $StageRoot 'from-finalizer.txt') -Value 'x' -Encoding ASCII
            $ManifestData['SetupFile'] = 'install.bat'
        }
        $data = @{ AppName = 'App'; SoftwareVersion = '1.0' }
        Write-StageManifest -Path (Join-Path $root 'stage-manifest.json') -ManifestData $data
        Should -Invoke Invoke-StageFinalization -ModuleName AppPackagerCommon -Times 1 -Exactly
        # A file the finalizer wrote must be covered by the recorded hashes.
        @($data['FileHashes']).RelativePath | Should -Contain 'from-finalizer.txt'
        $data['SetupFile'] | Should -Be 'install.bat'
    }

    It 'writes schema 4 with a plan digest and leaves legacy fields untouched' {
        $root = Join-Path $TestDrive 'legacy-shape'
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        Mock Invoke-StageFinalization -ModuleName AppPackagerCommon { }
        $path = Join-Path $root 'stage-manifest.json'
        $data = @{
            AppName = 'Legacy App'; Publisher = 'Vendor'; SoftwareVersion = '2.0'
            InstallerFile = 'setup.exe'; InstallerType = 'EXE'
            Detection = @{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\X'; Is64Bit = $true }
        }
        Write-StageManifest -Path $path -ManifestData $data
        $json = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $json.SchemaVersion | Should -Be 4
        $json.PlanDigest | Should -Match '^[0-9A-F]{64}$'
        # Everything a schema 3 consumer reads is unchanged.
        $json.AppName | Should -Be 'Legacy App'
        $json.Publisher | Should -Be 'Vendor'
        $json.SoftwareVersion | Should -Be '2.0'
        $json.InstallerFile | Should -Be 'setup.exe'
        $json.Detection.Type | Should -Be 'RegistryKey'
        $added = @(@($json.PSObject.Properties.Name) | Where-Object { @('AppName', 'Publisher', 'SoftwareVersion', 'InstallerFile', 'InstallerType', 'Detection', 'SchemaVersion', 'StagedAt', 'FileHashes') -notcontains $_ })
        $added | Should -Be @('PlanDigest')
    }

    It 'writes the build record after integrity verification' {
        $root = Join-Path $TestDrive 'build-record'
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        Mock Invoke-StageFinalization -ModuleName AppPackagerCommon { }
        Mock Write-BuildRecord -ModuleName AppPackagerCommon { }
        Write-StageManifest -Path (Join-Path $root 'stage-manifest.json') -ManifestData @{ AppName = 'A'; SoftwareVersion = '1' }
        Should -Invoke Write-BuildRecord -ModuleName AppPackagerCommon -Times 1 -Exactly
    }
}

Describe 'Get-StageManifestPlanDigest' {
    It 'ignores key order' {
        $a = [ordered]@{ AppName = 'A'; Detection = @{ Type = 'File'; FileName = 'x.exe' }; SoftwareVersion = '1' }
        $b = [ordered]@{ SoftwareVersion = '1'; Detection = @{ FileName = 'x.exe'; Type = 'File' }; AppName = 'A' }
        Get-StageManifestPlanDigest -ManifestData $a | Should -Be (Get-StageManifestPlanDigest -ManifestData $b)
    }

    It 'ignores FileHashes, PlanDigest and StagedAt' {
        $bare = @{ AppName = 'A' }
        $noisy = @{ AppName = 'A'; FileHashes = @(@{ RelativePath = 'a'; Sha256 = 'B'; Size = 1 }); PlanDigest = 'old'; StagedAt = (Get-Date -Format 'o') }
        Get-StageManifestPlanDigest -ManifestData $noisy | Should -Be (Get-StageManifestPlanDigest -ManifestData $bare)
    }

    It 'changes when a resolved command changes' {
        $before = Get-StageManifestPlanDigest -ManifestData @{ AppName = 'A'; InstallCommandLine = 'install.bat' }
        $after = Get-StageManifestPlanDigest -ManifestData @{ AppName = 'A'; InstallCommandLine = 'setup.exe /S' }
        $before | Should -Not -Be $after
    }

    It 'preserves array order' {
        $one = Get-StageManifestPlanDigest -ManifestData @{ DeploymentTypes = @('a', 'b') }
        $two = Get-StageManifestPlanDigest -ManifestData @{ DeploymentTypes = @('b', 'a') }
        $one | Should -Not -Be $two
    }
}

Describe 'Read-StageManifest schema range' {
    BeforeAll {
        function New-SchemaManifest {
            param($Version)
            $path = Join-Path $TestDrive ("schema-$Version.json")
            @{ SchemaVersion = $Version; AppName = 'A'; SoftwareVersion = '1'; FileHashes = @() } |
                ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding UTF8
            return $path
        }
    }

    It 'accepts schema 3 and schema 4' {
        (Read-StageManifest -Path (New-SchemaManifest 3)).SchemaVersion | Should -Be 3
        (Read-StageManifest -Path (New-SchemaManifest 4)).SchemaVersion | Should -Be 4
    }

    It 'refuses a newer schema with a clear message' {
        { Read-StageManifest -Path (New-SchemaManifest 5) } | Should -Throw '*newer than this AppPackager understands*'
    }

    It 'still applies the title mode' {
        $saved = $env:APP_PACKAGER_TITLE_MODE
        try {
            $env:APP_PACKAGER_TITLE_MODE = 'NoVersion'
            $path = Join-Path $TestDrive 'title-mode.json'
            @{ SchemaVersion = 4; AppName = 'Widget 1.0'; SoftwareVersion = '1.0'; FileHashes = @() } |
                ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding UTF8
            (Read-StageManifest -Path $path).AppName | Should -Be 'Widget'
        }
        finally {
            if ($saved) { $env:APP_PACKAGER_TITLE_MODE = $saved } else { Remove-Item Env:APP_PACKAGER_TITLE_MODE -ErrorAction SilentlyContinue }
        }
    }

    It 'lets the title mode recorded at Stage outrank the run-wide value' {
        $saved = $env:APP_PACKAGER_TITLE_MODE
        try {
            $env:APP_PACKAGER_TITLE_MODE = 'IncludeVersion'
            $recorded = Join-Path $TestDrive 'title-recorded.json'
            @{ SchemaVersion = 4; AppName = 'Widget'; SoftwareVersion = '1.0'; FileHashes = @(); TitleMode = 'NoVersion' } |
                ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $recorded -Encoding UTF8
            (Read-StageManifest -Path $recorded).AppName | Should -Be 'Widget'

            $inherited = Join-Path $TestDrive 'title-inherited.json'
            @{ SchemaVersion = 4; AppName = 'Widget'; SoftwareVersion = '1.0'; FileHashes = @() } |
                ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $inherited -Encoding UTF8
            (Read-StageManifest -Path $inherited).AppName | Should -Be 'Widget - 1.0'
        }
        finally {
            if ($saved) { $env:APP_PACKAGER_TITLE_MODE = $saved } else { Remove-Item Env:APP_PACKAGER_TITLE_MODE -ErrorAction SilentlyContinue }
        }
    }
}

# ---------------------------------------------------------------------------
# Timing
# ---------------------------------------------------------------------------

Describe 'Resolve-DeploymentTypeTiming' {
    It 'falls back to the parameters when the manifest carries no timing' {
        $r = Resolve-DeploymentTypeTiming -Timing $null -DefaultEstimated 15 -DefaultMaximum 30
        $r.EstimatedRuntimeMins | Should -Be 15
        $r.MaximumRuntimeMins | Should -Be 30
        $r.Source | Should -Be 'parameters'
    }

    It 'lets the manifest win over the parameters' {
        $r = Resolve-DeploymentTypeTiming -Timing @{ EstimatedMinutes = 5; MaximumMinutes = 90 } -DefaultEstimated 15 -DefaultMaximum 30
        $r.EstimatedRuntimeMins | Should -Be 5
        $r.MaximumRuntimeMins | Should -Be 90
        $r.Source | Should -Be 'manifest'
    }

    It 'inherits each value independently' {
        $r = Resolve-DeploymentTypeTiming -Timing @{ EstimatedMinutes = $null; MaximumMinutes = 60 } -DefaultEstimated 15 -DefaultMaximum 30
        $r.EstimatedRuntimeMins | Should -Be 15
        $r.MaximumRuntimeMins | Should -Be 60
    }

    It 'refuses an estimate above the maximum' {
        { Resolve-DeploymentTypeTiming -Timing @{ EstimatedMinutes = 90; MaximumMinutes = 30 } } | Should -Throw '*exceeds the maximum allowed run time*'
    }

    It 'refuses zero and out-of-range minutes' {
        { Resolve-DeploymentTypeTiming -Timing @{ MaximumMinutes = 0 } } | Should -Throw '*out of range*'
        { Resolve-DeploymentTypeTiming -Timing @{ MaximumMinutes = 2000 } } | Should -Throw '*out of range*'
    }

    It 'accepts a long maximum and reports it' {
        Mock Write-Log -ModuleName AppPackagerCommon { }
        (Resolve-DeploymentTypeTiming -Timing @{ EstimatedMinutes = 10; MaximumMinutes = 900 }).MaximumRuntimeMins | Should -Be 900
        Should -Invoke Write-Log -ModuleName AppPackagerCommon -ParameterFilter { $Level -eq 'WARN' } -Times 1 -Exactly
    }
}

Describe 'Get-ManifestDeploymentTypeSpecs timing' {
    It 'carries the base manifest timing onto a single deployment type' {
        $manifest = [pscustomobject]@{
            AppName = 'A'; Detection = [pscustomobject]@{ Type = 'RegistryKey' }
            Timing = [pscustomobject]@{ EstimatedMinutes = 7; MaximumMinutes = 70 }
        }
        $specs = @(Get-ManifestDeploymentTypeSpecs -Manifest $manifest -NetworkContentPath '\\srv\share' -AppName 'A')
        $specs[0].Timing.EstimatedMinutes | Should -Be 7
    }

    It 'lets a variant override the base timing' {
        $manifest = [pscustomobject]@{
            AppName = 'A'; Detection = [pscustomobject]@{ Type = 'RegistryKey' }
            Timing = [pscustomobject]@{ EstimatedMinutes = 7; MaximumMinutes = 70 }
            DeploymentTypes = @(
                [pscustomobject]@{ NameSuffix = 'x64'; Timing = [pscustomobject]@{ EstimatedMinutes = 20; MaximumMinutes = 40 } },
                [pscustomobject]@{ NameSuffix = 'x86' })
        }
        $specs = @(Get-ManifestDeploymentTypeSpecs -Manifest $manifest -NetworkContentPath '\\srv\share' -AppName 'A')
        $specs[0].Timing.EstimatedMinutes | Should -Be 20
        $specs[1].Timing.EstimatedMinutes | Should -Be 7
    }
}

# ---------------------------------------------------------------------------
# Deployment launchers
# ---------------------------------------------------------------------------

Describe 'Deployment launcher commands' {
    It 'writes the historical unsigned bat bodies byte for byte' {
        $out = Join-Path $TestDrive 'wrappers-unsigned'
        New-Item -ItemType Directory -Path $out -Force | Out-Null
        Write-ContentWrappers -OutputPath $out -InstallPs1Content 'exit 0' -UninstallPs1Content 'exit 0'
        $expected = (@(
            '@echo off',
            'PowerShell.exe -NonInteractive -ExecutionPolicy Bypass -File "%~dp0install.ps1"',
            'exit /b %ERRORLEVEL%'
        ) -join "`r`n")
        [IO.File]::ReadAllText((Join-Path $out 'install.bat')).TrimEnd("`r", "`n") | Should -Be $expected
    }

    It 'drops the execution-policy argument in signed deployment mode' {
        Mock Get-SigningPolicy -ModuleName AppPackagerCommon { [pscustomobject]@{ SignDeployment = $true } }
        $out = Join-Path $TestDrive 'wrappers-signed'
        New-Item -ItemType Directory -Path $out -Force | Out-Null
        Write-ContentWrappers -OutputPath $out -InstallPs1Content 'exit 0' -UninstallPs1Content 'exit 0' -InstallBatExitCode '3010'
        $bat = [IO.File]::ReadAllText((Join-Path $out 'install.bat'))
        $bat | Should -Not -Match 'ExecutionPolicy'
        $bat | Should -Match '\-NoProfile \-NonInteractive \-File'
        # The reboot exit-code contract survives the launcher change.
        $bat | Should -Match 'if %ERRORLEVEL% EQU 0 exit /b 3010'
    }

    It 'keeps the historical PSADT fallback command unchanged when unsigned' {
        $toolkit = Join-Path $TestDrive 'psadt-v3'
        New-Item -ItemType Directory -Path $toolkit -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $toolkit 'Deploy-Application.ps1') -Value '#' -Encoding ASCII
        $layout = Test-PsadtLayout -Path $toolkit
        $layout.InstallCommandLine | Should -Be 'powershell.exe -NonInteractive -ExecutionPolicy Bypass -File "Deploy-Application.ps1" -DeploymentType "Install"'
    }

    It 'removes bypass from the PSADT fallback in signed mode' {
        Mock Get-SigningPolicy -ModuleName AppPackagerCommon { [pscustomobject]@{ SignDeployment = $true } }
        $toolkit = Join-Path $TestDrive 'psadt-v4'
        New-Item -ItemType Directory -Path $toolkit -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $toolkit 'Invoke-AppDeployToolkit.ps1') -Value '#' -Encoding ASCII
        $layout = Test-PsadtLayout -Path $toolkit
        $layout.InstallCommandLine | Should -Not -Match 'ExecutionPolicy'
        $layout.InstallCommandLine | Should -Match 'Invoke-AppDeployToolkit\.ps1'
    }

    It 'rejects a launcher string that does not name the requested script' {
        Mock Write-Log -ModuleName AppPackagerCommon { }
        Mock New-DeploymentLauncherCommand -ModuleName AppPackagerCommon { [pscustomobject]@{ CommandLine = 'PowerShell.exe -File "other.ps1"' } }
        Get-DeploymentLauncherCommandLine -ScriptToken 'install.ps1' | Should -Match 'install\.ps1'
    }
}

# ---------------------------------------------------------------------------
# ConfigMgr detection script transport and read-back
# ---------------------------------------------------------------------------

Describe 'Resolve-DetectionScriptTransport' {
    It 'keeps ScriptText transport for a manifest with no script file' {
        $t = Resolve-DetectionScriptTransport -Detection ([pscustomobject]@{ ScriptText = 'exit 1' }) -ContentLocation $TestDrive
        $t.Transport | Should -Be 'Text'
        $t.Text | Should -Be 'exit 1'
    }

    It 'imports a finalized script file by path with its exact bytes' {
        $content = Join-Path $TestDrive 'content-a'
        New-Item -ItemType Directory -Path (Join-Path $content 'scripts') -Force | Out-Null
        $file = Join-Path $content 'scripts\detect.ps1'
        [IO.File]::WriteAllBytes($file, [Text.Encoding]::UTF8.GetBytes("Write-Output 'yes'`r`nexit 0"))
        $t = Resolve-DetectionScriptTransport -Detection ([pscustomobject]@{ ScriptFile = 'scripts\detect.ps1'; ScriptText = 'stale' }) -ContentLocation $content
        $t.Transport | Should -Be 'File'
        $t.Path | Should -Be $file
        $t.Bytes.Length | Should -Be ([IO.File]::ReadAllBytes($file)).Length
        $t.Text | Should -Match 'yes'
    }

    It 'refuses a detection script the content location does not carry' {
        $content = Join-Path $TestDrive 'content-b'
        New-Item -ItemType Directory -Path $content -Force | Out-Null
        { Resolve-DetectionScriptTransport -Detection ([pscustomobject]@{ ScriptFile = 'scripts\missing.ps1' }) -ContentLocation $content } |
            Should -Throw '*was not found under the content location*'
    }

    It 'refuses a finalized script above the site script size limit' {
        $content = Join-Path $TestDrive 'content-c'
        New-Item -ItemType Directory -Path $content -Force | Out-Null
        $file = Join-Path $content 'big.ps1'
        [IO.File]::WriteAllBytes($file, (New-Object byte[] ((Get-ConfigMgrDetectionScriptMaxBytes) + 1)))
        { Resolve-DetectionScriptTransport -Detection ([pscustomobject]@{ ScriptFile = 'big.ps1' }) -ContentLocation $content } |
            Should -Throw '*above the*'
    }
}

Describe 'Get-SdmPackageScriptText' {
    It 'reads the discovery script body' {
        $xml = "<AppMgmtDigest xmlns='http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest'><DeploymentType><Installer><CustomData><EnhancedDetectionMethod><Settings><DiscoveryScript><ScriptType>PowerShell</ScriptType><ScriptBody>exit 0</ScriptBody></DiscoveryScript></Settings></EnhancedDetectionMethod></CustomData></Installer></DeploymentType></AppMgmtDigest>"
        Get-SdmPackageScriptText -SdmPackageXml $xml | Should -Be 'exit 0'
    }

    It 'returns nothing for a document with no script body' {
        Get-SdmPackageScriptText -SdmPackageXml '<AppMgmtDigest><DeploymentType /></AppMgmtDigest>' | Should -BeNullOrEmpty
    }

    It 'returns nothing rather than throwing on unparseable content' {
        Mock Write-Log -ModuleName AppPackagerCommon { }
        Get-SdmPackageScriptText -SdmPackageXml 'not xml <' | Should -BeNullOrEmpty
    }
}

Describe 'Test-StoredDetectionScript' {
    It 'passes when the site stored the same script' {
        Mock Get-CMDeploymentTypeDetectionScript -ModuleName AppPackagerCommon { [pscustomobject]@{ Text = 'exit 0'; Bytes = [Text.Encoding]::UTF8.GetBytes('exit 0'); Encoded = $false } }
        $transport = [pscustomobject]@{ Transport = 'Text'; Text = 'exit 0'; Bytes = [Text.Encoding]::UTF8.GetBytes('exit 0'); TempPath = ''; Signed = $false }
        { Test-StoredDetectionScript -ApplicationName 'A' -DeploymentTypeName 'A' -Transport $transport } | Should -Not -Throw
    }

    It 'throws when the site stored different content' {
        Mock Get-CMDeploymentTypeDetectionScript -ModuleName AppPackagerCommon { [pscustomobject]@{ Text = 'exit 1'; Bytes = [Text.Encoding]::UTF8.GetBytes('exit 1'); Encoded = $false } }
        $transport = [pscustomobject]@{ Transport = 'Text'; Text = 'exit 0'; Bytes = [Text.Encoding]::UTF8.GetBytes('exit 0'); TempPath = ''; Signed = $false }
        { Test-StoredDetectionScript -ApplicationName 'A' -DeploymentTypeName 'A' -Transport $transport } |
            Should -Throw '*read back different detection script content*'
    }

    It 'throws when a signed detection cannot be read back at all' {
        Mock Get-CMDeploymentTypeDetectionScript -ModuleName AppPackagerCommon { $null }
        $transport = [pscustomobject]@{ Transport = 'File'; Text = 'exit 0'; Bytes = [Text.Encoding]::UTF8.GetBytes('exit 0'); TempPath = ''; Signed = $true }
        { Test-StoredDetectionScript -ApplicationName 'A' -DeploymentTypeName 'A' -Transport $transport } |
            Should -Throw '*cannot be verified*'
    }

    It 'removes the temporary import copy' {
        Mock Get-CMDeploymentTypeDetectionScript -ModuleName AppPackagerCommon { [pscustomobject]@{ Text = 'exit 0'; Bytes = [Text.Encoding]::UTF8.GetBytes('exit 0'); Encoded = $false } }
        $temp = Join-Path $TestDrive 'temp-import.ps1'
        Set-Content -LiteralPath $temp -Value 'exit 0' -Encoding ASCII
        $transport = [pscustomobject]@{ Transport = 'File'; Text = 'exit 0'; Bytes = [Text.Encoding]::UTF8.GetBytes('exit 0'); TempPath = $temp; Signed = $false }
        Test-StoredDetectionScript -ApplicationName 'A' -DeploymentTypeName 'A' -Transport $transport
        Test-Path -LiteralPath $temp | Should -BeFalse
    }
}

# ---------------------------------------------------------------------------
# Script global conditions: content identity and versioning
# ---------------------------------------------------------------------------

Describe 'Get-OrCreateGlobalConditionFromTemplate script identity' {
    InModuleScope AppPackagerCommon {
        BeforeAll {
            function Get-CMGlobalCondition { param($Name, $Id) }
            function New-CMGlobalConditionScript { param($Name, $DataType, $ScriptLanguage, $ScriptText, $FilePath, $Description) }
            function New-CMGlobalConditionWqlQuery { param($Name, $DataType, $Namespace, $Class, $Property, $Description) }
            $script:ScriptTemplate = [pscustomobject]@{
                Id = 'vpn'; Kind = 'Script'; GlobalConditionName = 'AppPackager VPN'
                DataType = 'Boolean'; Description = 'd'; RuleType = 'Boolean'
                ScriptText = @('exit 0')
            }
        }

        It 'reuses an existing condition whose script is unchanged' {
            Mock Get-CMGlobalCondition { [pscustomobject]@{ LocalizedDisplayName = 'AppPackager VPN'; CI_ID = 1 } }
            Mock Get-CMGlobalConditionScriptText { 'exit 0' }
            Mock New-CMGlobalConditionScript { throw 'must not create' }
            (Get-OrCreateGlobalConditionFromTemplate -Template $script:ScriptTemplate).CI_ID | Should -Be 1
        }

        It 'versions a changed script instead of mutating the shared condition' {
            Mock Get-CMGlobalCondition {
                if ($Name -eq 'AppPackager VPN') { return [pscustomobject]@{ LocalizedDisplayName = $Name; CI_ID = 1 } }
                return $null
            }
            # The existing condition carries different text; the new one must read
            # back exactly what was sent.
            Mock Get-CMGlobalConditionScriptText { if ($GlobalCondition.CI_ID -eq 1) { 'exit 1' } else { 'exit 0' } }
            Mock New-CMGlobalConditionScript { [pscustomobject]@{ LocalizedDisplayName = $Name; CI_ID = 2 } }

            $result = Get-OrCreateGlobalConditionFromTemplate -Template $script:ScriptTemplate
            $result.CI_ID | Should -Be 2
            Should -Invoke New-CMGlobalConditionScript -Times 1 -Exactly -ParameterFilter { $Name -match '^AppPackager VPN \([0-9a-f]{8}\)$' }
        }

        It 'throws when the site stores different script content than was sent' {
            Mock Get-CMGlobalCondition { $null }
            Mock New-CMGlobalConditionScript { [pscustomobject]@{ CI_ID = 3; SDMPackageXML = '<x><ScriptBody>tampered</ScriptBody></x>' } }
            { Get-OrCreateGlobalConditionFromTemplate -Template $script:ScriptTemplate } |
                Should -Throw '*read back different script content*'
        }

        It 'leaves a WQL condition on its existing name-match behavior' {
            Mock Get-CMGlobalCondition { [pscustomobject]@{ CI_ID = 9 } }
            $wql = [pscustomobject]@{ Id = 'w'; Kind = 'Wql'; GlobalConditionName = 'W'; DataType = 'String'; Namespace = 'n'; Class = 'c'; Property = 'p'; Description = 'd' }
            (Get-OrCreateGlobalConditionFromTemplate -Template $wql).CI_ID | Should -Be 9
        }

        It 'imports the exact signed bytes and verifies the stored copy against them' {
            $certificate = $null
            try {
                $certificate = New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=AppPackager Unit Test Condition' `
                    -CertStoreLocation 'Cert:\CurrentUser\My' -NotAfter (Get-Date).AddDays(1) -ErrorAction Stop
            }
            catch {
                Set-ItResult -Skipped -Because 'this host cannot create a code-signing certificate'
                return
            }
            try {
                $thumb = $certificate.Thumbprint.ToUpperInvariant()
                Mock Get-SigningPolicy {
                    [pscustomobject]@{
                        SignDetection = $false; SignRequirements = $true; SignDeployment = $false
                        RequireDetection = $false; RequireRequirements = $true; RequireDeployment = $false
                        CertificateThumbprint = $thumb; StoreLocation = 'CurrentUser'
                        TimestampServer = ''; TimestampRequired = $false; HashAlgorithm = 'SHA256'
                    }
                }
                Mock Get-CMGlobalCondition { $null }
                $script:SentBytes = $null
                Mock New-CMGlobalConditionScript {
                    $script:SentBytes = [System.IO.File]::ReadAllBytes($FilePath)
                    $body = "preview`r`n# ENCODEDSCRIPT # Begin Configuration Manager encoded script block # " +
                        [Convert]::ToBase64String($script:SentBytes) + " # ENCODEDSCRIPT# End Configuration Manager encoded script block"
                    [pscustomobject]@{
                        CI_ID = 4; LocalizedDisplayName = $Name
                        SDMPackageXML = ('<x><ScriptBody>' + [System.Security.SecurityElement]::Escape($body) + '</ScriptBody></x>')
                    }
                }

                $result = Get-OrCreateGlobalConditionFromTemplate -Template $script:ScriptTemplate
                $result.CI_ID | Should -Be 4
                Should -Invoke New-CMGlobalConditionScript -Times 1 -Exactly -ParameterFilter {
                    -not [string]::IsNullOrWhiteSpace($FilePath) -and [string]::IsNullOrWhiteSpace($ScriptText)
                }
                $script:SentBytes | Should -Not -BeNullOrEmpty
                # The import file carries the signed bytes as signed: no byte-order
                # mark was added or stripped between signing and import.
                $script:SentBytes[0] | Should -Not -Be 0xEF
                $verified = Test-ScriptSignatureBytes -Bytes $script:SentBytes
                $verified.SignatureIntact | Should -BeTrue
                $verified.Thumbprint | Should -Be $thumb
            }
            finally {
                Remove-Item -LiteralPath ("Cert:\CurrentUser\My\" + $certificate.Thumbprint) -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Intune adapter
# ---------------------------------------------------------------------------

Describe 'Get-IntuneAllowedArchitecture' {
    It 'maps the architectures Graph v1.0 documents' {
        Get-IntuneAllowedArchitecture -Architecture 'x64' | Should -Be 'x64'
        Get-IntuneAllowedArchitecture -Architecture 'AMD64' | Should -Be 'x64'
        Get-IntuneAllowedArchitecture -Architecture 'Intel' | Should -Be 'x86'
        Get-IntuneAllowedArchitecture -Architecture 'arm64' | Should -Be 'arm64'
    }

    It 'returns nothing for an architecture Graph has no value for' {
        Get-IntuneAllowedArchitecture -Architecture 'ia64' | Should -BeNullOrEmpty
    }
}

Describe 'Get-IntuneIdentityTag' {
    It 'uses the application and profile ids' {
        Get-IntuneIdentityTag -Manifest ([pscustomobject]@{ ApplicationId = 'catalog:package-7zip'; ProfileId = 'abc-1' }) |
            Should -Be 'AppPackager:catalog:package-7zip/abc-1'
    }

    It 'falls back to the default profile for a manifest staged before the workbench' {
        Get-IntuneIdentityTag -Manifest ([pscustomobject]@{ AppName = 'My App' }) | Should -Be 'AppPackager:legacy:My-App/default'
    }
}

Describe 'Get-IntuneScriptRuleSettings' {
    It 'enforces the signature check when the detection was signed' {
        $m = [pscustomobject]@{ ScriptSigning = [pscustomobject]@{ Detection = [pscustomobject]@{ Status = 'SignedAndVerified' } } }
        (Get-IntuneScriptRuleSettings -Manifest $m).EnforceSignatureCheck | Should -BeTrue
    }

    It 'leaves enforcement off for an unsigned detection under a permissive policy' {
        (Get-IntuneScriptRuleSettings -Manifest ([pscustomobject]@{ })).EnforceSignatureCheck | Should -BeFalse
    }

    It 'refuses to publish an unsigned detection when signatures are required' {
        Mock Get-SigningPolicy -ModuleName AppPackagerCommon { [pscustomobject]@{ RequireDetection = $true } }
        { Get-IntuneScriptRuleSettings -Manifest ([pscustomobject]@{ }) } | Should -Throw '*signatures are required*'
    }

    It 'takes the script host bitness from the definition' {
        $m = [pscustomobject]@{ Execution = [pscustomobject]@{ ScriptHost = 'x86' } }
        (Get-IntuneScriptRuleSettings -Manifest $m).RunAs32Bit | Should -BeTrue
    }
}

Describe 'Get-IntuneCompatibilityFindings' {
    It 'blocks a variant-split application' {
        $m = [pscustomobject]@{
            AppName = 'A'; Architecture = 'x64'
            Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'K'; Is64Bit = $true }
            DeploymentTypes = @([pscustomobject]@{ NameSuffix = 'x64' })
        }
        @(Get-IntuneCompatibilityFindings -Manifest $m | Where-Object { $_.Code -eq 'MultipleDeploymentTypes' }).Severity | Should -Be 'Blocking'
    }

    It 'blocks an OR-connected detection and names it' {
        $m = [pscustomobject]@{
            AppName = 'A'; Architecture = 'x64'
            Detection = [pscustomobject]@{ Type = 'Compound'; Connector = 'Or'; Clauses = @() }
        }
        $f = @(Get-IntuneCompatibilityFindings -Manifest $m | Where-Object { $_.Code -eq 'DetectionNotMappable' })
        $f.Count | Should -Be 1
        $f[0].Message | Should -Match 'OR-connected'
    }

    It 'records that runtime minutes are not applied in Intune' {
        $m = [pscustomobject]@{
            AppName = 'A'; Architecture = 'x64'; SetupFile = 'install.bat'
            Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'K'; Is64Bit = $true }
            Timing = [pscustomobject]@{ EstimatedMinutes = 10; MaximumMinutes = 60 }
        }
        @(Get-IntuneCompatibilityFindings -Manifest $m | Where-Object { $_.Code -eq 'TimingNotApplied' }).Severity | Should -Be 'Info'
    }

    It 'reports untranslated ConfigMgr requirements for review' {
        $m = [pscustomobject]@{
            AppName = 'A'; Architecture = 'x64'; SetupFile = 'install.bat'
            Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'K'; Is64Bit = $true }
            Requirements = @([pscustomobject]@{ ConditionId = 'vpn' })
        }
        @(Get-IntuneCompatibilityFindings -Manifest $m | Where-Object { $_.Code -eq 'RequirementsNotTranslated' }).Severity | Should -Be 'Review'
    }

    It 'passes a plain native manifest with no blocking finding' {
        $m = [pscustomobject]@{
            AppName = 'A'; Architecture = 'x64'; SetupFile = 'install.bat'
            Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'K'; Is64Bit = $true }
        }
        @(Get-IntuneCompatibilityFindings -Manifest $m | Where-Object { $_.Severity -eq 'Blocking' }).Count | Should -Be 0
    }
}

Describe 'Publish-IntuneWin32App request bodies' {
    BeforeAll {
        $script:IntuneManifest = [pscustomobject]@{
            AppName = 'Widget'; Publisher = 'Vendor'
            ApplicationId = 'catalog:package-widget'; ProfileId = 'managed-1'
            Architecture = 'arm64'
            SetupFile = 'toolkit\Invoke-AppDeployToolkit.exe'
            InstallCommandLine = 'Invoke-AppDeployToolkit.exe -DeploymentType Install'
            UninstallCommandLine = 'Invoke-AppDeployToolkit.exe -DeploymentType Uninstall'
            Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\Widget'; Is64Bit = $true }
        }
        function New-IntuneMocks {
            Mock Get-IntuneWinEncryptionInfo -ModuleName AppPackagerCommon {
                [pscustomobject]@{ SetupFile = 'install.bat'; FileName = 'w.intunewin'; UnencryptedContentSize = 10
                    EncryptionKey = 'k'; MacKey = 'm'; InitializationVector = 'i'; Mac = 'M'; ProfileIdentifier = 'p'; FileDigest = 'd'; FileDigestAlgorithm = 'SHA256' }
            }
            Mock Get-MsGraphToken -ModuleName AppPackagerCommon { 'token' }
            Mock Export-IntuneWinPayload -ModuleName AppPackagerCommon { [pscustomobject]@{ Path = (Join-Path $TestDrive 'payload.bin'); Size = 10 } }
            Mock Invoke-AzureBlobUpload -ModuleName AppPackagerCommon { }
            Mock Start-Sleep -ModuleName AppPackagerCommon { }
        }
    }

    BeforeEach {
        $script:GraphCalls = New-Object System.Collections.Generic.List[object]
        Set-Content -LiteralPath (Join-Path $TestDrive 'payload.bin') -Value 'x'
        Set-Content -LiteralPath (Join-Path $TestDrive 'w.intunewin') -Value 'x'
    }

    It 'sends allowedArchitectures, the real setup file, and the identity tag' {
        New-IntuneMocks
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
            $script:GraphCalls.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $Body })
            if ($Method -eq 'GET' -and $Uri -match 'mobileApps\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'GET') { return [pscustomobject]@{ azureStorageUri = 'https://blob.invalid/x'; uploadState = 'commitFileSuccess' } }
            return [pscustomobject]@{ id = 'app-1' }
        }

        Publish-IntuneWin32App -TenantId 't' -ClientId 'c' -ClientSecret 's' `
            -IntuneWinPath (Join-Path $TestDrive 'w.intunewin') -Manifest $script:IntuneManifest | Should -Be 'app-1'

        $create = @($script:GraphCalls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match 'mobileApps$' })[0]
        $create.Body['allowedArchitectures'] | Should -Be 'arm64'
        $create.Body.ContainsKey('applicableArchitectures') | Should -BeFalse
        $create.Body['setupFilePath'] | Should -Be 'toolkit\Invoke-AppDeployToolkit.exe'
        $create.Body['notes'] | Should -Be 'AppPackager:catalog:package-widget/managed-1'
        $create.Body['installCommandLine'] | Should -Be 'Invoke-AppDeployToolkit.exe -DeploymentType Install'
    }

    It 'updates the app carrying this identity tag rather than the first name match' {
        New-IntuneMocks
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
            if ($Method -eq 'GET' -and $Uri -match 'mobileApps\?') {
                return [pscustomobject]@{ value = @(
                    [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'other'; notes = 'AppPackager:catalog:package-widget/default' },
                    [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'mine'; notes = 'AppPackager:catalog:package-widget/managed-1' }) }
            }
            if ($Method -eq 'GET') { return [pscustomobject]@{ azureStorageUri = 'https://blob.invalid/x'; uploadState = 'commitFileSuccess' } }
            return [pscustomobject]@{ id = 'content-1' }
        }
        Publish-IntuneWin32App -TenantId 't' -ClientId 'c' -ClientSecret 's' `
            -IntuneWinPath (Join-Path $TestDrive 'w.intunewin') -Manifest $script:IntuneManifest | Should -Be 'mine'
    }

    It 'refuses an ambiguous name match with no identity tag' {
        New-IntuneMocks
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
            if ($Method -eq 'GET' -and $Uri -match 'mobileApps\?') {
                return [pscustomobject]@{ value = @(
                    [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'a'; notes = '' },
                    [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'b'; notes = '' }) }
            }
            return [pscustomobject]@{ id = 'x' }
        }
        { Publish-IntuneWin32App -TenantId 't' -ClientId 'c' -ClientSecret 's' `
            -IntuneWinPath (Join-Path $TestDrive 'w.intunewin') -Manifest $script:IntuneManifest } |
            Should -Throw '*none carries this application*'
    }

    It 'refuses to publish before any Graph call when a finding blocks the target' {
        New-IntuneMocks
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon { throw 'must not reach Graph' }
        $blocked = [pscustomobject]@{
            AppName = 'Widget'; Publisher = 'V'; Architecture = 'x64'
            Detection = [pscustomobject]@{ Type = 'Compound'; Connector = 'Or'; Clauses = @() }
        }
        { Publish-IntuneWin32App -TenantId 't' -ClientId 'c' -ClientSecret 's' `
            -IntuneWinPath (Join-Path $TestDrive 'w.intunewin') -Manifest $blocked } |
            Should -Throw '*cannot be published to Intune*'
    }

    It 'sends a signed detection script whose decoded bytes still verify' {
        $certificate = $null
        try {
            $certificate = New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=AppPackager Unit Test Signing' `
                -CertStoreLocation 'Cert:\CurrentUser\My' -NotAfter (Get-Date).AddDays(1) -ErrorAction Stop
        }
        catch {
            Set-ItResult -Skipped -Because 'this host cannot create a code-signing certificate'
            return
        }
        # Trust is an endpoint property: this asserts the signature stays intact through the
        # payload encoding, never that the build host trusts the signer.
        try {
            $scriptPath = Join-Path $TestDrive 'detect-signed.ps1'
            Set-Content -LiteralPath $scriptPath -Value "Write-Output 'Detected'`r`nexit 0" -Encoding UTF8
            Set-AuthenticodeSignature -FilePath $scriptPath -Certificate $certificate -HashAlgorithm SHA256 | Out-Null
            (Test-ScriptSignature -Path $scriptPath).SignatureIntact | Should -BeTrue

            $manifest = [pscustomobject]@{
                AppName = 'Signed Widget'; Publisher = 'V'; Architecture = 'x64'
                ContentRoot = "$TestDrive"
                Execution = [pscustomobject]@{ ScriptHost = 'x86' }
                ScriptSigning = [pscustomobject]@{ Detection = [pscustomobject]@{ Status = 'SignedAndVerified' } }
                Detection = [pscustomobject]@{ Type = 'Script'; ScriptFile = 'detect-signed.ps1'; ScriptText = 'stale text' }
            }
            $rules = @(ConvertTo-IntuneWin32Rules -Manifest $manifest)
            $rules[0]['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppPowerShellScriptRule'
            $rules[0].enforceSignatureCheck | Should -BeTrue
            $rules[0].runAs32Bit | Should -BeTrue

            # The Graph payload must carry the signed bytes, not the stale text.
            $decoded = [Convert]::FromBase64String($rules[0].scriptContent)
            $decoded | Should -Not -BeNullOrEmpty
            (Test-ScriptSignatureBytes -Bytes $decoded).SignatureIntact | Should -BeTrue
        }
        finally {
            if ($certificate) {
                Remove-Item -LiteralPath ("Cert:\CurrentUser\My\" + $certificate.Thumbprint) -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Shared download cache
# ---------------------------------------------------------------------------

Describe 'Invoke-CachedDownload shared cache' {
    It 'downloads once for two profile roots and delivers both files' {
        $base = Join-Path $TestDrive 'dl'
        New-Item -ItemType Directory -Path $base -Force | Out-Null
        Mock Invoke-DownloadWithRetry -ModuleName AppPackagerCommon {
            Set-Content -LiteralPath $OutFile -Value 'payload'
        }

        foreach ($profileId in @('p1', 'p2')) {
            $root = Join-Path (Join-Path $base 'profiles') $profileId
            New-Item -ItemType Directory -Path $root -Force | Out-Null
            Invoke-CachedDownload -Url 'https://example.invalid/setup.exe' -OutFile (Join-Path $root 'setup.exe') -Quiet -DownloadRoot $root
        }

        Get-Content -LiteralPath (Join-Path (Join-Path (Join-Path $base 'profiles') 'p1') 'setup.exe') | Should -Be 'payload'
        Get-Content -LiteralPath (Join-Path (Join-Path (Join-Path $base 'profiles') 'p2') 'setup.exe') | Should -Be 'payload'
        # Both profiles resolve to one cache entry rather than forking it.
        (Get-ChildItem -LiteralPath (Join-Path $base '_cache') -File).Count | Should -Be 1
    }

    It 'keeps the existing behavior when no download root is known' {
        Mock Invoke-DownloadWithRetry -ModuleName AppPackagerCommon { Set-Content -LiteralPath $OutFile -Value 'direct' }
        $target = Join-Path $TestDrive 'no-root.bin'
        Invoke-CachedDownload -Url 'https://example.invalid/y' -OutFile $target -Quiet -DownloadRoot ''
        Get-Content -LiteralPath $target | Should -Be 'direct'
    }
}


Describe 'ConvertFrom-CMEncodedScriptBody' {
    It 'decodes the encoded block the provider stores for a signed script' {
        $original = [Text.Encoding]::UTF8.GetBytes("Write-Output 'yes'`r`nexit 0`r\n# SIG # Begin signature block")
        $body = "Write-Output 'yes'`r`nexit 0`r`n# ENCODEDSCRIPT # Begin Configuration Manager encoded script block # " +
            [Convert]::ToBase64String($original) + " # ENCODEDSCRIPT# End Configuration Manager encoded script block"
        $decoded = ConvertFrom-CMEncodedScriptBody -Body $body
        $decoded.Encoded | Should -BeTrue
        [Convert]::ToBase64String($decoded.Bytes) | Should -Be ([Convert]::ToBase64String($original))
    }

    It 'passes an unsigned plain-text body through unchanged' {
        $decoded = ConvertFrom-CMEncodedScriptBody -Body 'exit 1'
        $decoded.Encoded | Should -BeFalse
        $decoded.Text | Should -Be 'exit 1'
    }
}

Describe 'Get-SdmPackageScriptText deployment type selection' {
    BeforeAll {
        $script:AppXml = "<AppMgmtDigest><Application /><DeploymentType><Title>App - x64</Title><Installer><CustomData><Args><Arg Name='ScriptBody'>x64 body</Arg></Args></CustomData></Installer></DeploymentType><DeploymentType><Title>App - x86</Title><Installer><CustomData><Args><Arg Name='ScriptBody'>x86 body</Arg></Args></CustomData></Installer></DeploymentType></AppMgmtDigest>"
    }

    It 'selects the body of the named deployment type, not the first one' {
        Get-SdmPackageScriptText -SdmPackageXml $script:AppXml -DeploymentTypeName 'App - x86' | Should -Be 'x86 body'
        Get-SdmPackageScriptText -SdmPackageXml $script:AppXml -DeploymentTypeName 'App - x64' | Should -Be 'x64 body'
    }

    It 'returns nothing rather than the wrong body when the title does not match' {
        Get-SdmPackageScriptText -SdmPackageXml $script:AppXml -DeploymentTypeName 'App - arm64' | Should -BeNullOrEmpty
    }

    It 'falls back to CustomData DetectionScript' {
        $xml = "<AppMgmtDigest><DeploymentType><Title>Only</Title><Installer><CustomData><DetectionScript>fallback body</DetectionScript></CustomData></Installer></DeploymentType></AppMgmtDigest>"
        Get-SdmPackageScriptText -SdmPackageXml $xml -DeploymentTypeName 'Only' | Should -Be 'fallback body'
    }
}

Describe 'Signed detection transport' {
    It 'refuses a signed detection that has no finalized script file' {
        { Resolve-DetectionScriptTransport -Detection ([pscustomobject]@{ ScriptText = 'exit 0' }) -ContentLocation $TestDrive -SignatureStatus 'SignedAndVerified' } |
            Should -Throw '*cannot be imported as script text*'
    }

    It 'compares signed content by bytes, not by normalized text' {
        Mock Get-CMDeploymentTypeDetectionScript -ModuleName AppPackagerCommon {
            [pscustomobject]@{ Text = 'exit 0'; Bytes = [Text.Encoding]::UTF8.GetBytes('exit 0 '); Encoded = $true }
        }
        Mock Test-ScriptSignatureBytes -ModuleName AppPackagerCommon { [pscustomobject]@{ Valid = $true; Status = 'Valid'; Thumbprint = 'T'; Reason = '' } }
        $transport = [pscustomobject]@{ Transport = 'File'; Text = 'exit 0'; Bytes = [Text.Encoding]::UTF8.GetBytes('exit 0'); TempPath = ''; Signed = $true }
        { Test-StoredDetectionScript -ApplicationName 'A' -DeploymentTypeName 'A' -Transport $transport } |
            Should -Throw '*different detection script bytes*'
    }

    It 'accepts a byte-identical signed read-back' {
        $bytes = [Text.Encoding]::UTF8.GetBytes("exit 0`r`n")
        Mock Get-CMDeploymentTypeDetectionScript -ModuleName AppPackagerCommon {
            [pscustomobject]@{ Text = 'exit 0'; Bytes = [Text.Encoding]::UTF8.GetBytes("exit 0`r`n"); Encoded = $true }
        }
        Mock Test-ScriptSignatureBytes -ModuleName AppPackagerCommon { [pscustomobject]@{ Valid = $true; Status = 'Valid'; Thumbprint = 'T'; Reason = '' } }
        $transport = [pscustomobject]@{ Transport = 'File'; Text = 'exit 0'; Bytes = $bytes; TempPath = ''; Signed = $true }
        { Test-StoredDetectionScript -ApplicationName 'A' -DeploymentTypeName 'A' -Transport $transport } | Should -Not -Throw
        Should -Invoke Test-ScriptSignatureBytes -ModuleName AppPackagerCommon -Times 1 -Exactly
    }

    It 'throws when the read-back signature does not verify' {
        Mock Get-CMDeploymentTypeDetectionScript -ModuleName AppPackagerCommon {
            [pscustomobject]@{ Text = 'exit 0'; Bytes = [Text.Encoding]::UTF8.GetBytes('exit 0'); Encoded = $true }
        }
        Mock Test-ScriptSignatureBytes -ModuleName AppPackagerCommon { [pscustomobject]@{ Valid = $false; Status = 'HashMismatch'; Thumbprint = ''; Reason = 'tampered' } }
        $transport = [pscustomobject]@{ Transport = 'File'; Text = 'exit 0'; Bytes = [Text.Encoding]::UTF8.GetBytes('exit 0'); TempPath = ''; Signed = $true }
        { Test-StoredDetectionScript -ApplicationName 'A' -DeploymentTypeName 'A' -Transport $transport } |
            Should -Throw '*did not verify after read-back*'
    }
}

Describe 'Operator command overrides reach the launcher check' {
    AfterEach { $env:APP_PACKAGER_COMMANDS = '' }

    It 'resolves the override into InstallCommandLine before the finalizer runs' {
        $root = Join-Path $TestDrive 'override-order'
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        $script:SeenInstall = 'not set'
        Mock Invoke-StageFinalization -ModuleName AppPackagerCommon {
            $script:SeenInstall = [string]$ManifestData['InstallCommandLine']
        }
        $env:APP_PACKAGER_COMMANDS = '{"SchemaVersion":1,"Install":"powershell.exe -ExecutionPolicy Bypass -File install.ps1","Uninstall":"uninstall.bat"}'
        $data = @{ AppName = 'App'; SoftwareVersion = '1.0' }
        Write-StageManifest -Path (Join-Path $root 'stage-manifest.json') -ManifestData $data
        $script:SeenInstall | Should -Be 'powershell.exe -ExecutionPolicy Bypass -File install.ps1'
        $data['CommandOverrides'].Install | Should -Be 'powershell.exe -ExecutionPolicy Bypass -File install.ps1'
    }
}

Describe 'Test-ResolvedDeploymentCommand' {
    BeforeAll {
        $script:BypassSpec = @([pscustomobject]@{
            InstallCommand   = 'powershell.exe -ExecutionPolicy Bypass -File install.ps1'
            UninstallCommand = 'uninstall.bat'
        })
        $script:CleanSpec = @([pscustomobject]@{ InstallCommand = 'install.bat'; UninstallCommand = 'uninstall.bat' })
    }

    It 'refuses a bypass command when the manifest reports a signed deployment category' {
        $manifest = [pscustomobject]@{
            ScriptSigning = [pscustomobject]@{ Deployment = [pscustomobject]@{ Status = 'SignedAndVerified' } }
        }
        { Test-ResolvedDeploymentCommand -Specs $script:BypassSpec -Manifest $manifest } |
            Should -Throw '*refuses this build*'
    }

    It 'refuses a bypass command when the policy requires deployment signatures' {
        $manifest = [pscustomobject]@{
            ScriptSigning = [pscustomobject]@{ Deployment = [pscustomobject]@{ Status = 'NotRequested' } }
        }
        $env:APP_PACKAGER_SIGNING = '{"SignDeployment":false,"RequireDeployment":true}'
        try {
            { Test-ResolvedDeploymentCommand -Specs $script:BypassSpec -Manifest $manifest } |
                Should -Throw '*refuses this build*'
        }
        finally { $env:APP_PACKAGER_SIGNING = '' }
    }

    It 'accepts a bypass-free command in signed mode and ignores an unsigned build' {
        $signed = [pscustomobject]@{
            ScriptSigning = [pscustomobject]@{ Deployment = [pscustomobject]@{ Status = 'SignedAndVerified' } }
        }
        (Test-ResolvedDeploymentCommand -Specs $script:CleanSpec -Manifest $signed).BypassFree | Should -BeTrue
        $unsigned = [pscustomobject]@{
            ScriptSigning = [pscustomobject]@{ Deployment = [pscustomobject]@{ Status = 'NotRequested' } }
        }
        Test-ResolvedDeploymentCommand -Specs $script:BypassSpec -Manifest $unsigned | Should -BeNullOrEmpty
    }
}

Describe 'Intune multi-deployment-type agreement' {
    It 'does not block a manifest whose DeploymentTypes collection is empty' {
        $m = [pscustomobject]@{
            AppName = 'A'; Architecture = 'x64'
            Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'K'; Is64Bit = $true }
            DeploymentTypes = @()
        }
        @(Get-IntuneCompatibilityFindings -Manifest $m | Where-Object { $_.Code -eq 'MultipleDeploymentTypes' }).Count | Should -Be 0
    }

    It 'blocks rather than throws when a variant entry cannot be resolved' {
        $m = [pscustomobject]@{
            AppName = 'A'; Architecture = 'x64'
            Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'K'; Is64Bit = $true }
            DeploymentTypes = @([pscustomobject]@{ ContentSubpath = 'x64' })
        }
        @(Get-IntuneCompatibilityFindings -Manifest $m | Where-Object { $_.Code -eq 'MultipleDeploymentTypes' }).Severity |
            Should -Be 'Blocking'
    }
}

Describe 'Intune script rule requires the finalized file' {
    It 'throws instead of re-encoding ScriptText when the named script file is missing' {
        $m = [pscustomobject]@{
            AppName = 'A'; Architecture = 'x64'; ContentRoot = "$TestDrive"
            Detection = [pscustomobject]@{ Type = 'Script'; ScriptFile = 'scripts\detect.ps1'; ScriptText = 'exit 0' }
        }
        { ConvertTo-IntuneWin32Rules -Manifest $m } | Should -Throw '*detect.ps1*'
    }

    It 'uses the finalized file when it exists' {
        $dir = Join-Path $TestDrive 'r16'
        New-Item -ItemType Directory -Path (Join-Path $dir 'scripts') -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $dir 'scripts\detect.ps1') -Value "Write-Output 'Detected'" -Encoding ASCII
        $m = [pscustomobject]@{
            AppName = 'A'; Architecture = 'x64'; ContentRoot = $dir
            Detection = [pscustomobject]@{ Type = 'Script'; ScriptFile = 'scripts\detect.ps1'; ScriptText = 'stale text' }
        }
        $rules = @(ConvertTo-IntuneWin32Rules -Manifest $m)
        $decoded = [Text.Encoding]::ASCII.GetString([Convert]::FromBase64String($rules[0].scriptContent))
        $decoded | Should -Match 'Detected'
        $decoded | Should -Not -Match 'stale text'
    }
}

Describe 'Intune existing-app lookup by identity tag' {
    BeforeAll {
        $script:R12Manifest = [pscustomobject]@{
            AppName = 'Widget'; Publisher = 'Vendor'
            ApplicationId = 'catalog:package-widget'; ProfileId = 'managed-1'
            Architecture = 'x64'
            SetupFile = 'install.bat'
            InstallCommandLine = 'install.bat'
            UninstallCommandLine = 'uninstall.bat'
            Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\Widget'; Is64Bit = $true }
        }
    }

    BeforeEach {
        $script:GraphCalls = New-Object System.Collections.Generic.List[object]
        Set-Content -LiteralPath (Join-Path $TestDrive 'payload.bin') -Value 'x'
        Set-Content -LiteralPath (Join-Path $TestDrive 'w.intunewin') -Value 'x'
        Mock Get-IntuneWinEncryptionInfo -ModuleName AppPackagerCommon {
            [pscustomobject]@{ SetupFile = 'install.bat'; FileName = 'w.intunewin'; UnencryptedContentSize = 10
                EncryptionKey = 'k'; MacKey = 'm'; InitializationVector = 'i'; Mac = 'M'; ProfileIdentifier = 'p'; FileDigest = 'd'; FileDigestAlgorithm = 'SHA256' }
        }
        Mock Get-MsGraphToken -ModuleName AppPackagerCommon { 'token' }
        Mock Export-IntuneWinPayload -ModuleName AppPackagerCommon { [pscustomobject]@{ Path = (Join-Path $TestDrive 'payload.bin'); Size = 10 } }
        Mock Invoke-AzureBlobUpload -ModuleName AppPackagerCommon { }
        Mock Start-Sleep -ModuleName AppPackagerCommon { }
    }

    It 'matches an app renamed in the tenant through its tag and never queries by display name' {
        $manifest = $script:R12Manifest
        $tag = Get-IntuneIdentityTag -Manifest $manifest
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
            $script:GraphCalls.Add([pscustomobject]@{ Method = $Method; Uri = $Uri })
            if ($Method -eq 'GET' -and $Uri -match 'isof') {
                return [pscustomobject]@{ value = @([pscustomobject]@{
                    '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'renamed-1'; displayName = 'Something Else'; notes = $tag }) }
            }
            if ($Method -eq 'GET' -and $Uri -match 'mobileApps\?') { return [pscustomobject]@{ value = @() } }
            if ($Method -eq 'GET') { return [pscustomobject]@{ azureStorageUri = 'https://blob.invalid/x'; uploadState = 'commitFileSuccess' } }
            return [pscustomobject]@{ id = 'renamed-1' }
        }

        Publish-IntuneWin32App -TenantId 't' -ClientId 'c' -ClientSecret 's' `
            -IntuneWinPath (Join-Path $TestDrive 'w.intunewin') -Manifest $manifest | Should -Be 'renamed-1'
        @($script:GraphCalls | Where-Object { $_.Uri -match "displayName eq" }).Count | Should -Be 0
        @($script:GraphCalls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match 'mobileApps$' }).Count | Should -Be 0
    }

    It 'refuses to publish when two apps carry the same identity tag' {
        $manifest = $script:R12Manifest
        $tag = Get-IntuneIdentityTag -Manifest $manifest
        Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
            if ($Method -eq 'GET' -and $Uri -match 'isof') {
                return [pscustomobject]@{ value = @(
                    [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'a'; notes = $tag },
                    [pscustomobject]@{ '@odata.type' = '#microsoft.graph.win32LobApp'; id = 'b'; notes = $tag }) }
            }
            throw 'must not reach any other Graph call'
        }
        { Publish-IntuneWin32App -TenantId 't' -ClientId 'c' -ClientSecret 's' `
            -IntuneWinPath (Join-Path $TestDrive 'w.intunewin') -Manifest $manifest } |
            Should -Throw '*identity tag*'
    }
}


Describe 'Read-back signature verification is integrity, not host trust' {
    It 'accepts an intact signature whose chain is not trusted on this host' {
        Mock Test-ScriptSignatureBytes -ModuleName AppPackagerCommon {
            [pscustomobject]@{ Valid = $false; SignatureIntact = $true; TrustedOnThisHost = $false
                Status = 'UnknownError'; Thumbprint = 'AABB'; Reason = 'root certificate which is not trusted' }
        }
        { Test-StoredScriptSignature -Bytes ([byte[]]@(1, 2, 3)) -Context 'deployment type ''X''' -ExpectedThumbprint 'AABB' } |
            Should -Not -Throw
    }

    It 'rejects a hash mismatch' {
        Mock Test-ScriptSignatureBytes -ModuleName AppPackagerCommon {
            [pscustomobject]@{ Valid = $false; SignatureIntact = $false; TrustedOnThisHost = $false
                Status = 'HashMismatch'; Thumbprint = 'AABB'; Reason = 'hash' }
        }
        { Test-StoredScriptSignature -Bytes ([byte[]]@(1)) -Context 'deployment type ''X''' } |
            Should -Throw '*did not verify after read-back*'
    }

    It 'rejects an unsigned read-back' {
        Mock Test-ScriptSignatureBytes -ModuleName AppPackagerCommon {
            [pscustomobject]@{ Valid = $false; SignatureIntact = $false; TrustedOnThisHost = $false
                Status = 'NotSigned'; Thumbprint = ''; Reason = 'no signature' }
        }
        { Test-StoredScriptSignature -Bytes ([byte[]]@(1)) -Context 'deployment type ''X''' } |
            Should -Throw '*did not verify after read-back*'
    }

    It 'rejects an intact signature from a different signer' {
        Mock Test-ScriptSignatureBytes -ModuleName AppPackagerCommon {
            [pscustomobject]@{ Valid = $true; SignatureIntact = $true; TrustedOnThisHost = $true
                Status = 'Valid'; Thumbprint = 'CCDD'; Reason = '' }
        }
        { Test-StoredScriptSignature -Bytes ([byte[]]@(1)) -Context 'deployment type ''X''' -ExpectedThumbprint 'AABB' } |
            Should -Throw '*not by the certificate this build signed with*'
    }
}

Describe 'Profile-scoped network content path' {
    AfterEach { $env:APP_PACKAGER_RUN_SNAPSHOT = '' }

    BeforeAll {
        function New-TestSnapshot {
            param($Id, $Name)
            $path = Join-Path $TestDrive ("snapshot-$Id.json")
            $body = @{ ProfileId = $Id; Profile = @{ Name = $Name } } | ConvertTo-Json -Depth 5
            Set-Content -LiteralPath $path -Value $body -Encoding ASCII
            return $path
        }
    }

    It 'keeps the default profile path unchanged' {
        $env:APP_PACKAGER_RUN_SNAPSHOT = New-TestSnapshot 'default' 'ignored'
        $root = Join-Path $TestDrive 'share-default'
        $p = Get-NetworkContentPath -FileServerPath $root -VendorFolder 'Vendor' -AppFolder 'App' -Version '1.0'
        $p | Should -Be (Join-Path (Join-Path (Join-Path (Join-Path $root 'Applications') 'Vendor') 'App') '1.0')
    }

    It 'scopes the version folder by profile name in both layouts' {
        $env:APP_PACKAGER_RUN_SNAPSHOT = New-TestSnapshot 'abcdef12-1' 'Managed Build!'
        $root = Join-Path $TestDrive 'share-named'
        $nested = Get-NetworkContentPath -FileServerPath $root -VendorFolder 'Vendor' -AppFolder 'App' -Version '1.0'
        (Split-Path -Leaf $nested) | Should -Be '1.0-ManagedBuild'
        $flat = Get-NetworkContentPath -FileServerPath $root -VendorFolder 'Vendor' -AppFolder 'App' -Version '1.0' -Layout Flat
        (Split-Path -Leaf $flat) | Should -Be 'Vendor-App-1.0-ManagedBuild'
    }

    It 'falls back to the first eight characters of the profile id when the profile has no name' {
        $env:APP_PACKAGER_RUN_SNAPSHOT = New-TestSnapshot 'abcdef1234-7' ''
        $root = Join-Path $TestDrive 'share-unnamed'
        $p = Get-NetworkContentPath -FileServerPath $root -VendorFolder 'Vendor' -AppFolder 'App' -Version '2.5'
        (Split-Path -Leaf $p) | Should -Be '2.5-abcdef12'
    }

    It 'gives two profiles of one version different folders' {
        $root = Join-Path $TestDrive 'share-two'
        $env:APP_PACKAGER_RUN_SNAPSHOT = New-TestSnapshot 'p-1' 'Alpha'
        $first = Get-NetworkContentPath -FileServerPath $root -VendorFolder 'Vendor' -AppFolder 'App' -Version '3.0'
        $env:APP_PACKAGER_RUN_SNAPSHOT = New-TestSnapshot 'p-2' 'Beta'
        $second = Get-NetworkContentPath -FileServerPath $root -VendorFolder 'Vendor' -AppFolder 'App' -Version '3.0'
        $first | Should -Not -Be $second
    }
}

Describe 'Network paths tolerate a trailing dot' {
    It 'trims trailing dots from the vendor and app segments' {
        $root = Join-Path $TestDrive 'share-dot'
        $p = Get-NetworkAppRoot -FileServerPath $root -VendorFolder 'KDE e.V.' -AppFolder 'KDiff3.'
        (Split-Path -Leaf $p) | Should -Be 'KDiff3'
        (Split-Path -Leaf (Split-Path -Parent $p)) | Should -Be 'KDE e.V'
        $p | Should -Not -Match '\.\|\.$'
    }

    It 'trims a trailing dot from the Flat layout folder name' {
        $root = Join-Path $TestDrive 'share-dot-flat'
        $p = Get-NetworkContentPath -FileServerPath $root -VendorFolder 'KDE e.V' -AppFolder 'KDiff3' -Version '1.12.4.' -Layout Flat
        (Split-Path -Leaf $p) | Should -Be 'KDE e.V-KDiff3-1.12.4'
    }
}
