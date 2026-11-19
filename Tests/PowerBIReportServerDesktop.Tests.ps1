#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $script:PackagerPath = Join-Path $PSScriptRoot '..\Packagers\package-powerbi-desktop-report-server.ps1'
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:PackagerPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Get-PowerBIReportServerDesktopVersion', 'Get-PowerBIReportServerDownloadInfo', 'Get-PowerBIReportServerInstallerVersion',
                      'New-PowerBIReportServerInstallWrapper', 'New-PowerBIReportServerUninstallWrapper', 'Invoke-StagePowerBIReportServerDesktop') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    function Write-Log { param($Message, $Level, [switch]$Quiet) }
    $script:ChangeLogUrl = 'https://example.invalid/changelog'
    $script:DownloadCenterUrl = 'https://example.invalid/download'
    $script:InstallerFileName = 'PBIDesktopSetupRS_x64.exe'

    # The shape of the change log page on Microsoft Learn: the report server first, then Power BI Desktop (optimized for Power BI Report Server).
    $script:ChangeLogHtml = @'
<h2 id="september-2026">September 2026</h2>
<h3 id="power-bi-report-server">Power BI Report Server</h3>
<ul>
<li><p><em>Version: 1.27.9818.31285 (build 15.0.1122.106), Released: TBD</em></p></li>
</ul>
<h3 id="power-bi-desktop-optimized-for-power-bi-report-server">Power BI Desktop (optimized for Power BI Report Server)</h3>
<ul>
<li><p><em>Version: 2.158.1602.0 (September 2026), Released: October 26, 2026</em></p>
<ul><li>Fixed issue with Oracle proxy user connections.</li></ul>
</li>
<li><p><em>Version: 2.158.1177.0 (September 2026), Released: September 29, 2026</em></p></li>
</ul>
<h2 id="may-2026">May 2026</h2>
<h3 id="power-bi-desktop-optimized-for-power-bi-report-server-1">Power BI Desktop (optimized for Power BI Report Server)</h3>
<ul><li><p><em>Version: 2.154.956.0 (May 2026), Released: May 21, 2026</em></p></li></ul>
'@
    $script:DownloadCenterHtml = '<html><script>window.__DLCDetails__={"dlcDetailsView":{"downloadFile":[' +
        '{"name":"PBIDesktopSetupRS_x64.exe","url":"https://download.microsoft.com/download/a/b/PBIDesktopSetupRS_x64.exe","version":"15.0.1122.106"},' +
        '{"name":"PowerBIReportServer.exe","url":"https://download.microsoft.com/download/a/b/PowerBIReportServer.exe","version":"15.0.1122.106"}]}}</script></html>'
    function Test-Parses { param([string]$Text) $tk = $null; $er = $null; [void][System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tk, [ref]$er); @($er).Count -eq 0 }
}

Describe 'Power BI Desktop (Report Server) version' {
    It 'reads the newest Desktop version, not the report server version' {
        function curl.exe { $global:LASTEXITCODE = 0; $script:ChangeLogHtml }
        Get-PowerBIReportServerDesktopVersion -Quiet | Should -Be '2.158.1602.0'
    }

    It 'never takes the report server version when the Desktop list of a section is empty' {
        function curl.exe { $global:LASTEXITCODE = 0; '<h3>Power BI Desktop (optimized for Power BI Report Server)</h3><p>No release.</p><h3>Power BI Report Server</h3><ul><li><p><em>Version: 1.27.9818.31285 (build 15.0.1122.106)</em></p></li></ul>' }
        Get-PowerBIReportServerDesktopVersion -Quiet | Should -BeNullOrEmpty
    }

    It 'returns null when the change log has no Desktop entry' {
        function curl.exe { $global:LASTEXITCODE = 0; '<h3>Power BI Report Server</h3><li>Version: 1.27.9818.31285</li>' }
        Get-PowerBIReportServerDesktopVersion -Quiet | Should -BeNullOrEmpty
    }

    It 'returns null when the page cannot be read' {
        function curl.exe { $global:LASTEXITCODE = 22; '' }
        Get-PowerBIReportServerDesktopVersion -Quiet | Should -BeNullOrEmpty
    }
}

Describe 'Power BI Desktop (Report Server) download' {
    It 'takes the version from the change log and the Desktop installer from the Download Center page' {
        function curl.exe {
            $global:LASTEXITCODE = 0
            if ("$args" -like '*changelog*') { $script:ChangeLogHtml } else { $script:DownloadCenterHtml }
        }
        $r = Get-PowerBIReportServerDownloadInfo -Quiet
        $r.Version | Should -Be '2.158.1602.0'
        $r.FileName | Should -Be 'PBIDesktopSetupRS_x64.exe'
        $r.DownloadUrl | Should -Be 'https://download.microsoft.com/download/a/b/PBIDesktopSetupRS_x64.exe'
    }

    It 'returns null when the page lists no Desktop installer' {
        function curl.exe {
            $global:LASTEXITCODE = 0
            if ("$args" -like '*changelog*') { $script:ChangeLogHtml } else { '<script>window.__DLCDetails__={"dlcDetailsView":{"downloadFile":[{"name":"PowerBIReportServer.exe","url":"u"}]}}</script>' }
        }
        Get-PowerBIReportServerDownloadInfo -Quiet | Should -BeNullOrEmpty
    }

    It 'returns null when the page has no metadata block' {
        function curl.exe {
            $global:LASTEXITCODE = 0
            if ("$args" -like '*changelog*') { $script:ChangeLogHtml } else { '<html></html>' }
        }
        Get-PowerBIReportServerDownloadInfo -Quiet | Should -BeNullOrEmpty
    }
}

Describe 'Power BI Desktop (Report Server) scripts' {
    It 'writes an install script that runs the installer silently with the documented options' {
        $text = New-PowerBIReportServerInstallWrapper -InstallerFile 'PBIDesktopSetupRS_x64.exe'
        Test-Parses $text | Should -BeTrue
        $text | Should -Match "Join-Path \`$PSScriptRoot 'PBIDesktopSetupRS_x64\.exe'"
        $text | Should -Match "'-quiet', '-norestart', 'ACCEPT_EULA=1', 'DISABLE_UPDATE_NOTIFICATION=1'"
        $text | Should -Match 'exit \$proc\.ExitCode'
    }

    It 'writes an uninstall script that runs the installer with -uninstall' {
        $text = New-PowerBIReportServerUninstallWrapper -InstallerFile 'PBIDesktopSetupRS_x64.exe'
        Test-Parses $text | Should -BeTrue
        $text | Should -Match "'-uninstall', '-quiet', '-norestart'"
    }
}

Describe 'Power BI Desktop (Report Server) stage' {
    BeforeAll {
        $script:BaseDownloadRoot = Join-Path $TestDrive 'PowerBIDesktopRS'
        $script:Info = [pscustomobject]@{ Version = '2.158.1602.0'; FileName = 'PBIDesktopSetupRS_x64.exe'; DownloadUrl = 'https://example.invalid/PBIDesktopSetupRS_x64.exe' }
    }
    BeforeEach {
        if (Test-Path -LiteralPath $script:BaseDownloadRoot) { Remove-Item -LiteralPath $script:BaseDownloadRoot -Recurse -Force }
        $script:Manifest = $null
        $script:Downloads = 0
        $script:VersionCalls = 0
        Mock Get-PowerBIReportServerDownloadInfo { $script:Info }
        Mock Invoke-DownloadWithRetry { $script:Downloads++; Set-Content -LiteralPath $OutFile -Value 'installer' }
        Mock Assert-InstallerSignature { 'Microsoft Corporation' }
        Mock Get-BurnBundleId { '{72b00526-f00d-4cb9-a561-7e8f4806a9be}' }
        Mock Get-PowerBIReportServerInstallerVersion { '2.158.1602.0' }
        Mock Write-ContentWrappers { }
        Mock Write-StageManifest { $script:Manifest = $ManifestData }
    }

    It 'stages the installer and detects the bundle entry by DisplayVersion in the 32-bit view' {
        $null = Invoke-StagePowerBIReportServerDesktop
        $d = $script:Manifest.Detection
        $d.Type | Should -Be 'RegistryKeyValue'
        $d.RegistryKeyRelative | Should -Be 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{72b00526-f00d-4cb9-a561-7e8f4806a9be}'
        $d.ValueName | Should -Be 'DisplayVersion'
        $d.Operator | Should -Be 'GreaterEquals'
        $d.ExpectedValue | Should -Be '2.158.1602.0'
        $d.Is64Bit | Should -BeFalse
        $script:Manifest.SoftwareVersion | Should -Be '2.158.1602.0'
        $script:Manifest.InstallArgs | Should -Be '-quiet -norestart ACCEPT_EULA=1 DISABLE_UPDATE_NOTIFICATION=1'
        $script:Manifest.UninstallArgs | Should -Be '-uninstall -quiet -norestart'
        $script:Manifest.RunningProcess | Should -Contain 'PBIDesktop'
        Test-Path -LiteralPath (Join-Path $script:BaseDownloadRoot '2.158.1602.0\PBIDesktopSetupRS_x64.exe') | Should -BeTrue
        (Get-Content -LiteralPath (Join-Path $script:BaseDownloadRoot 'staged-version.txt') -Raw).Trim() | Should -Be '2.158.1602.0'
        Should -Invoke Assert-InstallerSignature -Times 1 -ParameterFilter { $ExpectedSigner -eq 'Microsoft Corporation' -and $RemoveOnFailure }
    }

    It 'stops and stages nothing when the installer reports another version than the change log' {
        Mock Get-PowerBIReportServerInstallerVersion { '2.157.0.0' }
        { Invoke-StagePowerBIReportServerDesktop } | Should -Throw '*reports version 2.157.0.0*names version 2.158.1602.0*'
        $script:Manifest | Should -BeNullOrEmpty
        Test-Path -LiteralPath (Join-Path $script:BaseDownloadRoot 'PBIDesktopSetupRS_x64.exe') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:BaseDownloadRoot 'staged-version.txt') | Should -BeFalse
    }

    It 'downloads a cached installer of another version again' {
        New-Item -ItemType Directory -Path $script:BaseDownloadRoot -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:BaseDownloadRoot 'PBIDesktopSetupRS_x64.exe') -Value 'old'
        Mock Get-PowerBIReportServerInstallerVersion { $script:VersionCalls++; if ($script:VersionCalls -eq 1) { '2.100.0.0' } else { '2.158.1602.0' } }
        $null = Invoke-StagePowerBIReportServerDesktop
        $script:Downloads | Should -Be 1
        $script:Manifest.SoftwareVersion | Should -Be '2.158.1602.0'
    }

    It 'keeps a cached installer of the right version and does not download it' {
        New-Item -ItemType Directory -Path $script:BaseDownloadRoot -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:BaseDownloadRoot 'PBIDesktopSetupRS_x64.exe') -Value 'current'
        $null = Invoke-StagePowerBIReportServerDesktop
        $script:Downloads | Should -Be 0
    }

    It 'stages nothing when the signature check fails' {
        Mock Assert-InstallerSignature { throw 'The installer has no valid Authenticode signature. Nothing was staged.' }
        { Invoke-StagePowerBIReportServerDesktop } | Should -Throw '*Authenticode*'
        $script:Manifest | Should -BeNullOrEmpty
    }

    It 'stops when the version or the download cannot be resolved' {
        Mock Get-PowerBIReportServerDownloadInfo { $null }
        { Invoke-StagePowerBIReportServerDesktop } | Should -Throw '*Could not resolve*'
    }
}

Describe 'Power BI Desktop (Report Server) packager contract' {
    BeforeAll {
        $script:Text = Get-Content -LiteralPath $script:PackagerPath -Raw
        $script:StandardText = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\Packagers\package-powerbi-desktop.ps1') -Raw
    }

    It 'tags the header' {
        $script:Text | Should -Match '(?m)^Vendor: Microsoft\r?$'
        $script:Text | Should -Match '(?m)^CMName: Microsoft Power BI Desktop \(optimized for Power BI Report Server\)\r?$'
        $script:Text | Should -Match '(?m)^IconSource: External\r?$'
        $script:Text | Should -Match '(?m)^ArpArch: x64\r?$'
        $script:Text | Should -Match '(?m)^DownloadPageUrl: https://www\.microsoft\.com/download/details\.aspx\?id=105945\r?$'
        $script:Text | Should -Match '(?m)^ReleaseNotesUrl: https://learn\.microsoft\.com/power-bi/report-server/changelog\r?$'
    }

    It 'carries no CPE, because NVD has no entry for this application' {
        $script:Text | Should -Not -Match '(?m)^CPE:'
        $script:Text | Should -Match 'There is no CPE for this edition'
    }

    It 'has its own name, folder and entries, apart from the Power BI Desktop packager' {
        ([regex]::Match($script:Text, '(?m)^CMName: (.+?)\r?$').Groups[1].Value) | Should -Not -Be ([regex]::Match($script:StandardText, '(?m)^CMName: (.+?)\r?$').Groups[1].Value)
        $script:Text | Should -Match 'PowerBIDesktopRS'
        $script:Text | Should -Not -Match 'Microsoft Power BI Desktop\\bin'
    }

    It 'checks the signature of Microsoft Corporation and reads the bundle id from the installer' {
        $script:Text | Should -Match "Assert-InstallerSignature -Path \`$localInstaller -ExpectedSigner 'Microsoft Corporation' -RemoveOnFailure"
        $script:Text | Should -Match 'Get-BurnBundleId -Path \$localInstaller'
    }

    It 'builds no download URL by guessing' {
        $script:Text | Should -Not -Match 'download\.microsoft\.com/download/'
    }

    It 'names an Add or Remove Programs pattern that matches this edition and not the other one' {
        $pattern = ([regex]::Match($script:Text, '(?m)^ArpName: (.+?)\r?$').Groups[1].Value)
        'Microsoft Power BI Desktop (September 2026) (x64)' | Should -BeLike $pattern
        'Microsoft Power BI Desktop (May 2027) (x64)' | Should -BeLike $pattern
        'Microsoft Power BI Desktop (x64)' | Should -Not -BeLike $pattern
        'Microsoft PowerBI Desktop (x64)' | Should -Not -BeLike $pattern
    }

    It 'keeps the two Power BI Desktop packagers apart: each pattern matches its own entries and not the other edition''s' {
        $std = ([regex]::Match($script:StandardText, '(?m)^ArpName: (.+?)\r?$').Groups[1].Value)
        $rs = ([regex]::Match($script:Text, '(?m)^ArpName: (.+?)\r?$').Groups[1].Value)
        # The entries that a computer with both editions holds (read on a lab computer): the bundle entries of the 32-bit view, the Windows Installer entries of the 64-bit view.
        $standardEntries = @('Microsoft PowerBI Desktop (x64)', 'Microsoft Power BI Desktop (x64)')
        $reportServerEntries = @('Microsoft PowerBI Desktop (x64) (September 2026)', 'Microsoft Power BI Desktop (September 2026) (x64)')
        'Microsoft PowerBI Desktop (x64)' | Should -BeLike $std
        foreach ($name in $reportServerEntries) { $name | Should -Not -BeLike $std }
        'Microsoft Power BI Desktop (September 2026) (x64)' | Should -BeLike $rs
        foreach ($name in $standardEntries) { $name | Should -Not -BeLike $rs }
        $std | Should -Not -Be $rs
    }
}
