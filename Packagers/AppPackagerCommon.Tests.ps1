#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x tests for AppPackagerCommon shared module.

.DESCRIPTION
    Tests pure-logic and local-filesystem functions. Does NOT require MECM,
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
        $json.SchemaVersion | Should -Be 3
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
        $v2Manifest.SchemaVersion | Should -Be 3
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
        $manifest.SchemaVersion | Should -Be 3
        $manifest.PSObject.Properties.Name | Should -Contain 'FileHashes'
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
        }

        AfterAll {
            Remove-Item -Path function:\Get-CMApplication -ErrorAction SilentlyContinue
            Remove-Item -Path function:\Get-CMDeploymentType -ErrorAction SilentlyContinue
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
            Mock Get-CMApplication { [pscustomobject]@{ CI_ID = 1234 } }
            Mock Get-CMDeploymentType { [pscustomobject]@{ LocalizedDisplayName = 'Test App - 1.0' } }
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
            } | Should -Throw '*Multiple existing MECM applications*'
        }

        It 'fails before MECM app lookup when network content does not match manifest hashes' {
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
# calls Start-CMContentDistribution after creating the MECM Application.
#
# Testing this path cleanly requires either:
#   (a) refactoring the auto-distribute block into its own helper that takes
#       the prefs path as a parameter (then tested in isolation), or
#   (b) mocking the full MECM cmdlet surface (Connect-CMSite,
#       New-CMApplication, Get-CMApplication, Add-CMScriptDeploymentType,
#       Remove-CMApplicationRevisionHistoryByCIId, Start-CMContentDistribution)
#       and writing a temp AppPackager.preferences.json at the real repo root.
#
# Option (a) is the right design; option (b) would clobber dev prefs at repo
# root during test runs. Leaving as a TODO until the refactor lands.
