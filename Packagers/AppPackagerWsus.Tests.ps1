#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x tests for the AppPackagerWsus publisher.

.DESCRIPTION
    The WSUS administration API is not required. Every server call goes
    through the module's adapter functions, which the publishing tests
    replace with Pester mocks inside the module scope, so no WSUS server is
    contacted and nothing is published.

.EXAMPLE
    Invoke-Pester .\AppPackagerWsus.Tests.ps1
#>

BeforeAll {
    Import-Module "$PSScriptRoot\AppPackagerWsus.psd1" -Force

    $script:Root = Join-Path ([System.IO.Path]::GetTempPath()) ('apwsus-' + [guid]::NewGuid().ToString('N').Substring(0, 10))
    New-Item -ItemType Directory -Path $script:Root -Force | Out-Null
    $script:Content = Join-Path $script:Root 'content'
    New-Item -ItemType Directory -Path $script:Content -Force | Out-Null
    foreach ($name in @('setup.exe', 'custom.mst', 'product.msi', 'install.bat', 'stage-manifest.json', 'app-icon.ico')) {
        [System.IO.File]::WriteAllText((Join-Path $script:Content $name), $name)
    }
    $script:StandardInstallScript = @(
        '$exePath = Join-Path $PSScriptRoot ''setup.exe'''
        '$proc = Start-Process -FilePath $exePath -ArgumentList ''/S'' -Wait -PassThru -NoNewWindow'
        '$proc.ExitCode | Out-File (Join-Path $PSScriptRoot ''exitcode.txt'') -NoNewline'
        'exit $proc.ExitCode'
    ) -join "`r`n"
    [System.IO.File]::WriteAllText((Join-Path $script:Content 'install.ps1'), $script:StandardInstallScript)

    function New-TestManifest {
        param([hashtable]$Overrides = @{})
        $data = [ordered]@{
            SchemaVersion   = 4
            AppName         = 'Contoso Tool 5.4'
            Publisher       = 'Contoso'
            SoftwareVersion = '5.4.2'
            InstallerFile   = 'setup.exe'
            InstallerType   = 'EXE'
            InstallArgs     = '/S'
            Architecture    = 'x64'
            ApplicationId   = 'catalog:package-contosotool'
            ProfileId       = 'default'
            Detection       = [pscustomobject]@{
                Type                = 'RegistryKeyValue'
                RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ContosoTool'
                ValueName           = 'DisplayVersion'
                DisplayVersion      = '5.4.2'
                Is64Bit             = $true
            }
        }
        foreach ($key in $Overrides.Keys) { $data[$key] = $Overrides[$key] }
        return [pscustomobject]$data
    }

    function Test-WellFormedRule {
        param([string]$Rule)
        $wrapped = '<r xmlns:bar="urn:bar" xmlns:lar="urn:lar" xmlns:msiar="urn:msiar">' + $Rule + '</r>'
        try { [void]([xml]$wrapped); return $true } catch { return $false }
    }
}

AfterAll {
    Remove-Item -LiteralPath $script:Root -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'ConvertTo-WsusPublishSettings' {
    It 'fills every default from an empty input' {
        $s = ConvertTo-WsusPublishSettings -InputObject @{}
        $s.ServerName | Should -Be ''
        $s.PortNumber | Should -Be 8530
        $s.UseSsl | Should -BeFalse
        $s.ContainsKey('PackageType') | Should -BeFalse
        $s.Classification | Should -Be 'Updates'
        $s.ApprovalGroup | Should -Be ''
        $s.DeclineSuperseded | Should -BeFalse
    }

    It 'defaults the port to 8531 when SSL is on and no port is given' {
        (ConvertTo-WsusPublishSettings -InputObject @{ ServerName = 'wsus01'; UseSsl = $true }).PortNumber | Should -Be 8531
    }

    It 'keeps an explicit port and trims the server name' {
        $s = ConvertTo-WsusPublishSettings -InputObject ([pscustomobject]@{ ServerName = ' wsus01.contoso.com '; PortNumber = '443'; UseSsl = $true })
        $s.ServerName | Should -Be 'wsus01.contoso.com'
        $s.PortNumber | Should -Be 443
    }

    It 'refuses a server name that is not a host name' {
        { ConvertTo-WsusPublishSettings -InputObject @{ ServerName = 'wsus01;calc' } } | Should -Throw '*not a host name*'
    }

    It 'refuses an unknown classification' {
        { ConvertTo-WsusPublishSettings -InputObject @{ Classification = 'Drivers' } } | Should -Throw '*classification*'
    }

    It 'refuses the retired package type Application and any type other than Update' {
        { ConvertTo-WsusPublishSettings -InputObject @{ PackageType = 'Application' } } | Should -Throw '*installed products only*Remove the PackageType setting*'
        { ConvertTo-WsusPublishSettings -InputObject @{ PackageType = 'Bundle' } } | Should -Throw '*package type*'
        { ConvertTo-WsusPublishSettings -InputObject @{ PackageType = 'Update' } } | Should -Not -Throw
    }
}

Describe 'Identity line' {
    It 'round-trips the identity and version through an update description' {
        $tag = Get-WsusIdentityTag -Manifest (New-TestManifest)
        $tag | Should -Be 'AppPackager:catalog:package-contosotool/default'
        $description = "Contoso Tool 5.4.2`r`n`r`n" + (Get-WsusIdentityLine -IdentityTag $tag -PackageType Update -Version '2026.08.2-200') + "`r`n"
        $parsed = Get-WsusUpdateIdentity -Description $description
        $parsed.IdentityTag | Should -Be $tag
        $parsed.PackageType | Should -Be 'Update'
        $parsed.Version | Should -Be '2026.08.2-200'
    }

    It 'reads an identity line that carries no version' {
        $parsed = Get-WsusUpdateIdentity -Description ("x`r`n" + (Get-WsusIdentityLine -IdentityTag 'AppPackager:catalog:package-a/default' -PackageType Application))
        $parsed.PackageType | Should -Be 'Application'
        $parsed.Version | Should -Be ''
    }

    It 'derives a legacy key from the publisher and the title for a manifest without workbench ids' {
        $manifest = New-TestManifest -Overrides @{ ApplicationId = $null; ProfileId = $null }
        Get-WsusIdentityTag -Manifest $manifest | Should -Be 'AppPackager:legacy:Contoso:Contoso-Tool-5.4/default'
    }

    It 'keeps two publishers apart when their legacy titles match' {
        $a = New-TestManifest -Overrides @{ ApplicationId = $null; ProfileId = $null; AppName = 'Contoso Helper'; Publisher = 'Acme'; SoftwareVersion = '1.0' }
        $b = New-TestManifest -Overrides @{ ApplicationId = $null; ProfileId = $null; AppName = 'Contoso Helper'; Publisher = 'Northwind'; SoftwareVersion = '1.0' }
        $tagA = Get-WsusIdentityTag -Manifest $a
        $tagB = Get-WsusIdentityTag -Manifest $b
        $tagA | Should -Not -Be $tagB
        Get-WsusUpdateIdentity -Description (Get-WsusIdentityLine -IdentityTag $tagA -PackageType Update -Version '1.0') | ForEach-Object IdentityTag | Should -Be $tagA
    }

    It 'keeps one legacy identity across the releases of one major when the title carries the version' {
        $a = New-TestManifest -Overrides @{ ApplicationId = $null; ProfileId = $null; AppName = '7-Zip 26.03 (x64 edition)'; SoftwareVersion = '26.03' }
        $b = New-TestManifest -Overrides @{ ApplicationId = $null; ProfileId = $null; AppName = '7-Zip 26.04 (x64 edition)'; SoftwareVersion = '26.04' }
        Get-WsusIdentityTag -Manifest $a | Should -Be 'AppPackager:legacy:Contoso:7-Zip-26-x64-edition/default'
        Get-WsusIdentityTag -Manifest $b | Should -Be (Get-WsusIdentityTag -Manifest $a)
    }

    It 'keeps side-by-side majors apart when the title differs only by the version' {
        $net8 = New-TestManifest -Overrides @{ ApplicationId = $null; ProfileId = $null; AppName = 'Microsoft .NET 8.0.31 - Windows Server Hosting'; Publisher = 'Microsoft Corporation'; SoftwareVersion = '8.0.31' }
        $net8Older = New-TestManifest -Overrides @{ ApplicationId = $null; ProfileId = $null; AppName = 'Microsoft .NET 8.0.30 - Windows Server Hosting'; Publisher = 'Microsoft Corporation'; SoftwareVersion = '8.0.30' }
        $net10 = New-TestManifest -Overrides @{ ApplicationId = $null; ProfileId = $null; AppName = 'Microsoft .NET 10.0.12 - Windows Server Hosting'; Publisher = 'Microsoft Corporation'; SoftwareVersion = '10.0.12' }
        Get-WsusIdentityTag -Manifest $net8 | Should -Be 'AppPackager:legacy:Microsoft-Corporation:Microsoft-.NET-8-Windows-Server-Hosting/default'
        Get-WsusIdentityTag -Manifest $net8Older | Should -Be (Get-WsusIdentityTag -Manifest $net8)
        Get-WsusIdentityTag -Manifest $net10 | Should -Not -Be (Get-WsusIdentityTag -Manifest $net8)
    }

    It 'returns nothing for a description without an identity line' {
        Get-WsusUpdateIdentity -Description 'Published by another tool' | Should -BeNullOrEmpty
        Get-WsusUpdateIdentity -Description '' | Should -BeNullOrEmpty
    }
}

Describe 'ConvertFrom-WsusCatalogInput' {
    It 'extracts bare IDs and catalog links once each, in input order' {
        $text = @"
12345678-90ab-cdef-1234-567890abcdef
https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=0e3c9a1b-2f4d-4c5e-8a6b-7d8e9f0a1b2c
12345678-90AB-CDEF-1234-567890ABCDEF, 0e3c9a1b-2f4d-4c5e-8a6b-7d8e9f0a1b2c
"@
        $ids = @(ConvertFrom-WsusCatalogInput -Text $text)
        $ids.Count | Should -Be 2
        $ids[0] | Should -Be ([guid]'12345678-90ab-cdef-1234-567890abcdef')
        $ids[1] | Should -Be ([guid]'0e3c9a1b-2f4d-4c5e-8a6b-7d8e9f0a1b2c')
    }

    It 'returns nothing for text without an update ID' {
        @(ConvertFrom-WsusCatalogInput -Text 'KB5034441').Count | Should -Be 0
        @(ConvertFrom-WsusCatalogInput -Text '').Count | Should -Be 0
    }
}

Describe 'ConvertTo-WsusMsiCommandLine' {
    It 'keeps property assignments and drops msiexec switches and log files' {
        ConvertTo-WsusMsiCommandLine -InstallArgs '/qn TRANSFORMS=custom.mst /l*v "C:\log dir\x.log" ALLUSERS=1 PROP="a b"' |
            Should -Be 'TRANSFORMS=custom.mst ALLUSERS=1 PROP="a b"'
    }

    It 'turns /norestart into REBOOT=ReallySuppress once' {
        ConvertTo-WsusMsiCommandLine -InstallArgs '/qn /norestart' | Should -Be 'REBOOT=ReallySuppress'
    }

    It 'keeps an explicit REBOOT property instead of adding one' {
        ConvertTo-WsusMsiCommandLine -InstallArgs '/qn /norestart REBOOT=Force' | Should -Be 'REBOOT=Force'
    }

    It 'returns an empty string for switches only' {
        ConvertTo-WsusMsiCommandLine -InstallArgs '/qn' | Should -Be ''
    }
}

Describe 'Payload files' {
    It 'carries every recorded stage file with the installer first, never a generated file' {
        $recorded = @('app-icon.ico', 'Data1.cab', 'install.bat', 'install.ps1', 'scripts\detect.ps1', 'setup.ini', 'setup.exe', 'uninstall.ps1', 'tool.png', 'Patches/fix.msp')
        $manifest = New-TestManifest -Overrides @{ Icon = 'tool.png'; FileHashes = @($recorded | ForEach-Object { [pscustomobject]@{ RelativePath = $_; Sha256 = ('0' * 64); Size = 1 } }) }
        $files = InModuleScope AppPackagerWsus -Parameters @{ M = $manifest } {
            param($M)
            Get-WsusPayloadFiles -Manifest $M
        }
        @($files) | Should -Be @('setup.exe', 'Data1.cab', 'setup.ini', 'Patches\fix.msp')
    }
}

Describe 'Payload paths' {
    It 'resolves <Path> inside the content folder' -TestCases @(
        @{ Path = 'setup.exe' }
        @{ Path = 'Data\x.cab' }
    ) {
        param($Path)
        $full = InModuleScope AppPackagerWsus -Parameters @{ C = $script:Content; P = $Path } { param($C, $P) Resolve-WsusPayloadPath -ContentFolder $C -RelativePath $P }
        $full | Should -Be ([System.IO.Path]::GetFullPath((Join-Path $script:Content $Path)))
    }

    It 'refuses <Path>' -TestCases @(
        @{ Path = '..\outside.txt' }
        @{ Path = 'Data\..\..\outside.txt' }
        @{ Path = '..' }
        @{ Path = 'C:outside.txt' }
        @{ Path = 'C:\Windows\win.ini' }
        @{ Path = '\Windows\win.ini' }
        @{ Path = '\\server\share\setup.exe' }
        @{ Path = 'setup.exe:hidden' }
        @{ Path = '' }
    ) {
        param($Path)
        { InModuleScope AppPackagerWsus -Parameters @{ C = $script:Content; P = $Path } { param($C, $P) Resolve-WsusPayloadPath -ContentFolder $C -RelativePath $P } } |
            Should -Throw '*content folder*'
    }

    It 'blocks a manifest that records a file outside the content folder' {
        $outside = Join-Path $script:Root 'outside.txt'
        [System.IO.File]::WriteAllText($outside, 'marker')
        $hash = (Get-FileHash -LiteralPath $outside -Algorithm SHA256).Hash
        $manifest = New-TestManifest -Overrides @{ FileHashes = @(
                [pscustomobject]@{ RelativePath = 'setup.exe'; Sha256 = (Get-FileHash -LiteralPath (Join-Path $script:Content 'setup.exe') -Algorithm SHA256).Hash; Size = 9 }
                [pscustomobject]@{ RelativePath = '..\outside.txt'; Sha256 = $hash; Size = 6 }
            ) }
        $finding = @(Get-WsusCompatibilityFindings -Manifest $manifest -ContentFolder $script:Content | Where-Object Code -eq 'PayloadOutsideStage')
        $finding.Count | Should -Be 1
        $finding[0].Severity | Should -Be 'Blocking'
        $finding[0].Message | Should -Match '\.\.\\outside\.txt'
        { InModuleScope AppPackagerWsus -Parameters @{ M = $manifest; C = $script:Content } { param($M, $C) Assert-WsusPayloadIntegrity -Manifest $M -ContentFolder $C -Files @('setup.exe', '..\outside.txt') } } |
            Should -Throw '*outside the content folder*'
    }

    It 'reports an installer outside the content folder once' {
        $findings = @(Get-WsusCompatibilityFindings -Manifest (New-TestManifest -Overrides @{ InstallerFile = '..\setup.exe' }) -ContentFolder $script:Content)
        @($findings | Where-Object Code -eq 'PayloadOutsideStage').Count | Should -Be 1
        @($findings | Where-Object { $_.Code -in 'InstallerMissing', 'InstallerInSubfolder' }).Count | Should -Be 0
    }
}

Describe 'Install script check' {
    It 'accepts <Name>' -TestCases @(
        @{ Name = 'the standard wrapper'; Text = '' }
        @{ Name = 'an MSI wrapper that closes the running copy'; Text = @'
Get-Process tool -ErrorAction SilentlyContinue | Stop-Process -Force
$msi = Join-Path $PSScriptRoot 'product.msi'
$p = Start-Process msiexec.exe -ArgumentList @('/i', "`"$msi`"", '/qn') -Wait -PassThru
exit $p.ExitCode
'@ }
        @{ Name = 'a direct installer call'; Text = @'
& (Join-Path $PSScriptRoot 'setup.exe') /S
exit $LASTEXITCODE
'@ }
        @{ Name = 'a wrapper written with aliases'; Text = @'
gps tool -ErrorAction SilentlyContinue | kill -Force
sleep 2
$p = saps (Join-Path $PSScriptRoot 'setup.exe') -ArgumentList '/S' -Wait -PassThru
exit $p.ExitCode
'@ }
        @{ Name = 'a wrapper that waits on its process object'; Text = @'
$p = Start-Process (Join-Path $PSScriptRoot 'setup.exe') -ArgumentList '/S' -PassThru
$p.WaitForExit()
if ([string]$p.ExitCode.ToString().Trim() -eq '0') { exit 0 }
exit $p.ExitCode
'@ }
        @{ Name = 'a wrapper that finds its MSI and reports errors'; Text = @'
$msi = Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.msi' | Select-Object -First 1
if (-not $msi) { Write-Error 'No MSI in the content folder.'; exit 1 }
$p = Start-Process msiexec.exe -ArgumentList @('/i', $msi.FullName, '/qn') -Wait -PassThru
exit $p.ExitCode
'@ }
    ) {
        param($Name, $Text)
        if (-not $Text) { $Text = $script:StandardInstallScript }
        $path = Join-Path $script:Root ('install-{0}.ps1' -f [guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($path, $Text)
        @(InModuleScope AppPackagerWsus -Parameters @{ P = $path } { param($P) Get-WsusInstallScriptExtras -Path $P }).Count | Should -Be 0
    }

    It 'names <Name>' -TestCases @(
        @{ Name = 'a machine environment variable'; Expected = '[Environment]::SetEnvironmentVariable'; Text = @'
$p = Start-Process setup.exe -ArgumentList '/S' -Wait -PassThru
[Environment]::SetEnvironmentVariable('DBEAVER_AI_DISABLED', 'true', 'Machine')
exit 0
'@ }
        @{ Name = 'a second process'; Expected = '2 process launches'; Text = @'
Start-Process setup.exe -ArgumentList '/x' -Wait
Start-Process (Join-Path $env:TEMP 'x\setup.exe') -Wait
'@ }
        @{ Name = 'a registry write'; Expected = 'Set-ItemProperty'; Text = @'
Start-Process setup.exe -Wait
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Contoso' -Name AutoUpdate -Value 0
'@ }
        @{ Name = 'a file copy'; Expected = 'Copy-Item'; Text = @'
Start-Process setup.exe -Wait
Copy-Item (Join-Path $PSScriptRoot 'settings.xml') 'C:\ProgramData\Contoso'
'@ }
        @{ Name = 'a script that does not parse'; Expected = 'install.ps1 does not parse'; Text = 'Start-Process setup.exe -Wait {' }
        @{ Name = 'a file deleted through its own method'; Expected = '.Delete()'; Text = @'
$p = Start-Process setup.exe -ArgumentList '/S' -Wait -PassThru
(Get-Item 'C:\ProgramData\Contoso\settings.xml').Delete()
exit $p.ExitCode
'@ }
        @{ Name = 'a file changed through a property'; Expected = '$f.IsReadOnly assignment'; Text = @'
Start-Process setup.exe -Wait
$f = Get-Item 'C:\ProgramData\Contoso\settings.xml'
$f.IsReadOnly = $true
'@ }
        @{ Name = 'an environment variable for the installer'; Expected = '$env:CONTOSO_SILENT assignment'; Text = @'
$env:CONTOSO_SILENT = '1'
Start-Process setup.exe -Wait
'@ }
    ) {
        param($Name, $Text, $Expected)
        $path = Join-Path $script:Root ('install-{0}.ps1' -f [guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($path, $Text)
        @(InModuleScope AppPackagerWsus -Parameters @{ P = $path } { param($P) Get-WsusInstallScriptExtras -Path $P }) | Should -Contain $Expected
    }
}

Describe 'Compare-WsusVersion' {
    It 'orders <Left> against <Right> as <Expected>' -TestCases @(
        @{ Left = '5.4.2'; Right = '5.4.3'; Expected = -1 }
        @{ Left = '5.10'; Right = '5.9'; Expected = 1 }
        @{ Left = '5.4'; Right = '5.4.0'; Expected = 0 }
        @{ Left = '2026.08.2-200'; Right = '2026.08.2-199'; Expected = 1 }
        @{ Left = '24.08'; Right = '24.8'; Expected = 0 }
        @{ Left = '1.0.100000000000000000000'; Right = '1.0.99999999999999999999'; Expected = 1 }
    ) {
        param($Left, $Right, $Expected)
        InModuleScope AppPackagerWsus -Parameters @{ L = $Left; R = $Right } { param($L, $R) Compare-WsusVersion -Left $L -Right $R } | Should -Be $Expected
    }

    It 'cannot compare a version without digits' {
        InModuleScope AppPackagerWsus { Compare-WsusVersion -Left 'latest' -Right '5.4' } | Should -BeNullOrEmpty
        InModuleScope AppPackagerWsus { Compare-WsusVersion -Left '' -Right '5.4' } | Should -BeNullOrEmpty
    }
}

Describe 'ConvertTo-WsusApplicabilityRules' {
    BeforeAll {
        # Evaluates the rule subset that the publisher writes against a
        # described computer. A registry version with fewer than four parts
        # is padded with zeros. A missing or unreadable value or file compares
        # as version 0.0.0.0, the way RegSzToVersion treats a missing value, so
        # an unguarded LessThan check holds on a computer without the product.
        function ConvertTo-TestVersion {
            param([AllowNull()][AllowEmptyString()][string]$Text)
            if ([string]::IsNullOrWhiteSpace($Text) -or $Text.Trim() -notmatch '^\d+(\.\d+){0,3}$') { return $null }
            $parts = @($Text.Trim().Split('.') | ForEach-Object { [int]$_ })
            while ($parts.Count -lt 4) { $parts += 0 }
            return New-Object System.Version ($parts[0], $parts[1], $parts[2], $parts[3])
        }

        function Test-TestComparison {
            param([AllowNull()][AllowEmptyString()][string]$Actual, [string]$Comparison, [string]$Expected)
            $left = ConvertTo-TestVersion $Actual
            if ($null -eq $left) { $left = New-Object System.Version (0, 0, 0, 0) }
            $order = $left.CompareTo((ConvertTo-TestVersion $Expected))
            switch ($Comparison) {
                'LessThan' { return $order -lt 0 }
                'LessThanOrEqualTo' { return $order -le 0 }
                'EqualTo' { return $order -eq 0 }
                'GreaterThanOrEqualTo' { return $order -ge 0 }
                'GreaterThan' { return $order -gt 0 }
                default { throw ('Unknown comparison ' + $Comparison) }
            }
        }

        function Test-TestNode {
            param($Node, [hashtable]$Computer, [string]$LoopTarget = '')
            $view = if ($Node.GetAttribute('RegType32') -eq 'true') { '32' } else { '64' }
            $subkey = $Node.GetAttribute('Subkey')
            # Inside a RegKeyLoop, Subkey "\" under HKEY_LOOP_TARGET is the
            # looped subkey itself; any other form never matches.
            if ($Node.GetAttribute('Key') -eq 'HKEY_LOOP_TARGET') {
                if (-not $LoopTarget) { throw 'HKEY_LOOP_TARGET outside a RegKeyLoop' }
                $subkey = if ($subkey -eq '\') { $LoopTarget } else { '|no-such-key' }
            }
            $registryKey = '{0}|{1}' -f $subkey, $view
            $fileKey = '{0}|{1}' -f $Node.GetAttribute('Csidl'), $Node.GetAttribute('Path')
            switch ($Node.LocalName) {
                'True' { return $true }
                'And' {
                    foreach ($child in $Node.ChildNodes) { if (-not (Test-TestNode -Node $child -Computer $Computer -LoopTarget $LoopTarget)) { return $false } }
                    return $true
                }
                'Or' {
                    foreach ($child in $Node.ChildNodes) { if (Test-TestNode -Node $child -Computer $Computer -LoopTarget $LoopTarget) { return $true } }
                    return $false
                }
                'RegKeyLoop' {
                    $parent = $subkey + '\'
                    $children = @($Computer.Registry.Keys | Where-Object { $_.EndsWith('|' + $view) } |
                        ForEach-Object { $_.Substring(0, $_.Length - $view.Length - 1) } |
                        Where-Object { $_.StartsWith($parent, [StringComparison]::OrdinalIgnoreCase) -and -not $_.Substring($parent.Length).Contains('\') })
                    $rule = @($Node.ChildNodes | Where-Object { $_.NodeType -eq 'Element' })[0]
                    $results = @($children | ForEach-Object { Test-TestNode -Node $rule -Computer $Computer -LoopTarget $_ })
                    switch ($Node.GetAttribute('TrueIf')) {
                        'Any' { return ($results -contains $true) }
                        'All' { return ($results -notcontains $false) }
                        'None' { return ($results -notcontains $true) }
                    }
                }
                'RegSz' {
                    $actual = if ($Computer.Registry.ContainsKey($registryKey)) { $Computer.Registry[$registryKey][$Node.GetAttribute('Value')] } else { $null }
                    if ($null -eq $actual) { return $false }
                    $data = $Node.GetAttribute('Data')
                    switch ($Node.GetAttribute('Comparison')) {
                        'EqualTo' { return ([string]$actual -ceq $data) }
                        'BeginsWith' { return ([string]$actual).StartsWith($data, [StringComparison]::Ordinal) }
                        'EndsWith' { return ([string]$actual).EndsWith($data, [StringComparison]::Ordinal) }
                        'Contains' { return ([string]$actual).Contains($data) }
                    }
                }
                'RegDword' {
                    $actual = if ($Computer.Registry.ContainsKey($registryKey)) { $Computer.Registry[$registryKey][$Node.GetAttribute('Value')] } else { $null }
                    if ($null -eq $actual -or $actual -isnot [int]) { return $false }
                    switch ($Node.GetAttribute('Comparison')) {
                        'EqualTo' { return ($actual -eq [int]$Node.GetAttribute('Data')) }
                        default { throw ('The test evaluator has no RegDword case for ' + $Node.GetAttribute('Comparison')) }
                    }
                }
                'RegKeyExists' { return $Computer.Registry.ContainsKey($registryKey) }
                'RegValueExists' { return ($Computer.Registry.ContainsKey($registryKey) -and $Computer.Registry[$registryKey].ContainsKey($Node.GetAttribute('Value'))) }
                'RegSzToVersion' {
                    $actual = if ($Computer.Registry.ContainsKey($registryKey)) { $Computer.Registry[$registryKey][$Node.GetAttribute('Value')] } else { $null }
                    return (Test-TestComparison -Actual $actual -Comparison $Node.GetAttribute('Comparison') -Expected $Node.GetAttribute('Data'))
                }
                'FileExists' { return $Computer.Files.ContainsKey($fileKey) }
                'FileVersion' {
                    $actual = if ($Computer.Files.ContainsKey($fileKey)) { $Computer.Files[$fileKey] } else { $null }
                    return (Test-TestComparison -Actual $actual -Comparison $Node.GetAttribute('Comparison') -Expected $Node.GetAttribute('Version'))
                }
                default { throw ('The test evaluator has no case for ' + $Node.LocalName) }
            }
        }

        # The Windows Update Agent offers an update that is not installed and
        # is installable.
        function Get-TestOutcome {
            param([Parameter(Mandatory)]$Rules, [hashtable]$Registry = @{}, [hashtable]$Files = @{})
            $computer = @{ Registry = $Registry; Files = $Files }
            $wrap = { param($Rule) [xml]('<r xmlns:bar="urn:bar" xmlns:lar="urn:lar">' + $Rule + '</r>') }
            $isInstalled = Test-TestNode -Node (& $wrap $Rules.IsInstalled).DocumentElement.FirstChild -Computer $computer
            $isInstallable = Test-TestNode -Node (& $wrap $Rules.IsInstallable).DocumentElement.FirstChild -Computer $computer
            return [pscustomobject]@{ Installed = $isInstalled; Offered = (-not $isInstalled -and $isInstallable) }
        }

        $script:ArpKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ContosoTool'
        $script:ProductKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{23170F69-40C1-2702-2408-000001000000}'
        $script:HostingBundleArp = [pscustomobject]@{
            Type              = 'ArpEntry'
            View              = '32'
            DisplayNamePrefix = 'Microsoft .NET 8.0.'
            DisplayNameSuffix = ' - Windows Server Hosting'
            Publisher         = 'Microsoft Corporation'
            Version           = '8.0.31'
        }
        # Add/Remove Programs entries as "key|view" registry rows.
        function New-TestArpRegistry {
            param([object[]]$Entries)
            $registry = @{}
            foreach ($e in $Entries) {
                $values = @{ DisplayName = $e.Name; Publisher = $(if ($e.Publisher) { $e.Publisher } else { 'Microsoft Corporation' }) }
                if ($e.Version) { $values.DisplayVersion = $e.Version }
                if ($null -ne $e.WindowsInstaller) { $values.WindowsInstaller = [int]$e.WindowsInstaller }
                $registry[('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{0}|{1}' -f $e.Key, $(if ($e.View) { $e.View } else { '32' }))] = $values
            }
            return $registry
        }
    }

    It 'maps a DisplayVersion detection to "this version or newer" and "present with an older version"' {
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest)
        $rules.IsInstalled | Should -Be ('<bar:RegSzToVersion Key="HKEY_LOCAL_MACHINE" Subkey="{0}" Value="DisplayVersion" Comparison="GreaterThanOrEqualTo" Data="5.4.2.0" />' -f $script:ArpKey)
        $rules.IsInstallable | Should -Be ('<lar:And><bar:RegValueExists Key="HKEY_LOCAL_MACHINE" Subkey="{0}" Value="DisplayVersion" Type="REG_SZ" /><bar:RegSzToVersion Key="HKEY_LOCAL_MACHINE" Subkey="{0}" Value="DisplayVersion" Comparison="LessThan" Data="5.4.2.0" /></lar:And>' -f $script:ArpKey)
    }

    It 'guards the older check, because a version comparison alone holds on a computer without the value' {
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest)
        $versionOnly = [regex]::Match($rules.IsInstallable, '<bar:RegSzToVersion [^>]*/>').Value
        $node = ([xml]('<r xmlns:bar="urn:bar" xmlns:lar="urn:lar">' + $versionOnly + '</r>')).DocumentElement.FirstChild
        Test-TestNode -Node $node -Computer @{ Registry = @{}; Files = @{} } | Should -BeTrue
        (Get-TestOutcome -Rules $rules).Offered | Should -BeFalse
    }

    It 'offers a registry-detected update only over an older installed version: <Name>' -TestCases @(
        @{ Name = 'product absent'; Values = $null; Installed = $false; Offered = $false }
        @{ Name = 'uninstall key without a version'; Values = @{}; Installed = $false; Offered = $false }
        @{ Name = 'older version'; Values = @{ DisplayVersion = '5.4.1' }; Installed = $false; Offered = $true }
        @{ Name = 'older version with two parts'; Values = @{ DisplayVersion = '5.3' }; Installed = $false; Offered = $true }
        @{ Name = 'target version'; Values = @{ DisplayVersion = '5.4.2' }; Installed = $true; Offered = $false }
        @{ Name = 'newer version'; Values = @{ DisplayVersion = '5.10' }; Installed = $true; Offered = $false }
    ) {
        param($Name, $Values, $Installed, $Offered)
        $registry = @{}
        if ($null -ne $Values) { $registry[('{0}|64' -f $script:ArpKey)] = $Values }
        $outcome = Get-TestOutcome -Rules (ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest)) -Registry $registry
        $outcome.Installed | Should -Be $Installed
        $outcome.Offered | Should -Be $Offered
    }

    It 'offers a file-version update only over an older installed version: <Name>' -TestCases @(
        @{ Name = 'product absent'; Version = $null; Installed = $false; Offered = $false }
        @{ Name = 'older version'; Version = '5.4.1.9'; Installed = $false; Offered = $true }
        @{ Name = 'target version'; Version = '5.4.2.0'; Installed = $true; Offered = $false }
        @{ Name = 'newer version'; Version = '6.0.0.0'; Installed = $true; Offered = $false }
    ) {
        param($Name, $Version, $Installed, $Offered)
        $detection = [pscustomobject]@{ Type = 'File'; FilePath = 'C:\Program Files\Contoso'; FileName = 'tool.exe'; PropertyType = 'Version'; Operator = 'GreaterEquals'; ExpectedValue = '5.4.2'; Is64Bit = $true }
        $files = @{}
        if ($Version) { $files['38|Contoso\tool.exe'] = $Version }
        $outcome = Get-TestOutcome -Rules (ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })) -Files $files
        $outcome.Installed | Should -Be $Installed
        $outcome.Offered | Should -Be $Offered
    }

    It 'requires the presence clauses and an older version for an AND detection: <Name>' -TestCases @(
        @{ Name = 'older version without the marker file'; Marker = $false; Version = '5.4.1'; Installed = $false; Offered = $false }
        @{ Name = 'marker file without the product'; Marker = $true; Version = $null; Installed = $false; Offered = $false }
        @{ Name = 'marker file and older version'; Marker = $true; Version = '5.4.1'; Installed = $false; Offered = $true }
        @{ Name = 'marker file and target version'; Marker = $true; Version = '5.4.2'; Installed = $true; Offered = $false }
    ) {
        param($Name, $Marker, $Version, $Installed, $Offered)
        $detection = [pscustomobject]@{ Type = 'Compound'; Connector = 'And'; Clauses = @(
                [pscustomobject]@{ Type = 'File'; FilePath = 'C:\Program Files\Contoso'; FileName = 'tool.exe'; PropertyType = 'Existence'; Is64Bit = $true }
                [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:ArpKey; ValueName = 'DisplayVersion'; DisplayVersion = '5.4.2'; Is64Bit = $true }) }
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })
        $registry = @{}
        if ($Version) { $registry[('{0}|64' -f $script:ArpKey)] = @{ DisplayVersion = $Version } }
        $files = @{}
        if ($Marker) { $files['38|Contoso\tool.exe'] = '' }
        $outcome = Get-TestOutcome -Rules $rules -Registry $registry -Files $files
        $outcome.Installed | Should -Be $Installed
        $outcome.Offered | Should -Be $Offered
        Test-WellFormedRule -Rule $rules.IsInstallable | Should -BeTrue
    }

    It 'requires every version clause of an AND detection to find its value: <Name>' -TestCases @(
        @{ Name = 'registry value missing, file older'; Registry = $null; File = '5.4.1.0'; Installed = $false; Offered = $false }
        @{ Name = 'both older'; Registry = '5.4.1'; File = '5.4.1.0'; Installed = $false; Offered = $true }
        @{ Name = 'registry older, file current'; Registry = '5.4.1'; File = '5.4.2.0'; Installed = $false; Offered = $true }
        @{ Name = 'both current'; Registry = '5.4.2'; File = '5.4.2.0'; Installed = $true; Offered = $false }
    ) {
        param($Name, $Registry, $File, $Installed, $Offered)
        $detection = [pscustomobject]@{ Type = 'Compound'; Connector = 'And'; Clauses = @(
                [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:ArpKey; ValueName = 'DisplayVersion'; DisplayVersion = '5.4.2'; Is64Bit = $true }
                [pscustomobject]@{ Type = 'File'; FilePath = 'C:\Program Files\Contoso'; FileName = 'tool.exe'; PropertyType = 'Version'; ExpectedValue = '5.4.2'; Is64Bit = $true }) }
        $registryState = @{}
        if ($Registry) { $registryState[('{0}|64' -f $script:ArpKey)] = @{ DisplayVersion = $Registry } }
        $outcome = Get-TestOutcome -Rules (ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })) -Registry $registryState -Files @{ '38|Contoso\tool.exe' = $File }
        $outcome.Installed | Should -Be $Installed
        $outcome.Offered | Should -Be $Offered
    }

    It 'offers an OR detection where one view is older and no view is current: <Name>' -TestCases @(
        @{ Name = 'absent in both views'; View64 = $null; View32 = $null; Installed = $false; Offered = $false }
        @{ Name = 'older in the 64-bit view'; View64 = '5.4.1'; View32 = $null; Installed = $false; Offered = $true }
        @{ Name = 'older in one view, current in the other'; View64 = '5.4.1'; View32 = '5.4.2'; Installed = $true; Offered = $false }
    ) {
        param($Name, $View64, $View32, $Installed, $Offered)
        $detection = [pscustomobject]@{ Type = 'Compound'; Connector = 'Or'; Clauses = @(
                [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:ArpKey; ValueName = 'DisplayVersion'; DisplayVersion = '5.4.2'; Is64Bit = $true }
                [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:ArpKey; ValueName = 'DisplayVersion'; DisplayVersion = '5.4.2'; Is64Bit = $false }) }
        $registry = @{}
        if ($View64) { $registry[('{0}|64' -f $script:ArpKey)] = @{ DisplayVersion = $View64 } }
        if ($View32) { $registry[('{0}|32' -f $script:ArpKey)] = @{ DisplayVersion = $View32 } }
        $outcome = Get-TestOutcome -Rules (ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })) -Registry $registry
        $outcome.Installed | Should -Be $Installed
        $outcome.Offered | Should -Be $Offered
    }

    It 'maps GreaterThan to an older check of LessThanOrEqualTo' {
        $detection = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:ArpKey; ValueName = 'DisplayVersion'; ExpectedValue = '5.4.2'; Operator = 'GreaterThan'; Is64Bit = $true }
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })
        $rules.IsInstalled | Should -Match 'Comparison="GreaterThan" Data="5\.4\.2\.0"'
        $rules.IsInstallable | Should -Match 'Comparison="LessThanOrEqualTo" Data="5\.4\.2\.0"'
    }

    It 'refuses a detection that cannot show an older installed version: <Name>' -TestCases @(
        @{ Name = 'file existence'; Message = '*only checks that a file or registry key exists*'; Detection = [pscustomobject]@{ Type = 'File'; FilePath = 'C:\Program Files\Contoso'; FileName = 'tool.exe'; PropertyType = 'Existence'; Is64Bit = $true } }
        @{ Name = 'folder named for this version'; Message = '*only checks that a file or registry key exists*'; Detection = [pscustomobject]@{ Type = 'File'; FilePath = 'C:\Program Files\Contoso\5.4.2'; FileName = 'tool.exe'; PropertyType = 'Existence'; Is64Bit = $true } }
        @{ Name = 'registry key existence'; Message = '*only checks that a file or registry key exists*'; Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\Contoso\Tool'; Is64Bit = $true } }
        @{ Name = 'OR with an existence clause'; Message = '*joins its clauses with OR*'; Detection = [pscustomobject]@{ Type = 'Compound'; Connector = 'Or'; Clauses = @(
                    [pscustomobject]@{ Type = 'File'; FilePath = 'C:\Program Files\Contoso'; FileName = 'tool.exe'; PropertyType = 'Existence'; Is64Bit = $true }
                    [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Contoso\Tool'; DisplayVersion = '5.4.2'; Is64Bit = $true }) } }
        @{ Name = 'value that is not a version'; Message = '*not a version of up to four numbers*'; Detection = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Contoso'; ValueName = 'Edition'; ExpectedValue = 'Pro & Team'; Operator = 'IsEquals'; Is64Bit = $true } }
        @{ Name = 'text prefix comparison'; Message = '*as text*'; Detection = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Contoso'; ValueName = 'DisplayVersion'; ExpectedValue = '5.4'; Operator = 'BeginsWith'; Is64Bit = $true } }
        @{ Name = 'version under a Windows Installer product key'; Message = '*Windows Installer product key*'; Detection = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{23170F69-40C1-2702-2408-000001000000}'; DisplayVersion = '24.08.00.0'; Is64Bit = $true } }
        @{ Name = 'Windows Installer product key existence'; Message = '*Windows Installer product key*'; Detection = [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{23170F69-40C1-2702-2408-000001000000}'; Is64Bit = $true } }
    ) {
        param($Name, $Message, $Detection)
        $manifest = New-TestManifest -Overrides @{ Detection = $Detection }
        { ConvertTo-WsusApplicabilityRules -Manifest $manifest } | Should -Throw $Message
        $finding = @(Get-WsusCompatibilityFindings -Manifest $manifest -ContentFolder $script:Content | Where-Object Code -eq 'DetectionNotMappable')
        $finding.Count | Should -Be 1
        $finding[0].Severity | Should -Be 'Blocking'
    }

    It 'names the change that makes an existence-only detection publishable, not a new-install fallback' {
        $detection = [pscustomobject]@{ Type = 'File'; FilePath = 'C:\Program Files\Contoso'; FileName = 'tool.exe'; PropertyType = 'Existence'; Is64Bit = $true }
        $message = ''
        try { [void](ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })) } catch { $message = $_.Exception.Message }
        $message | Should -Match 'Add a file version or a registry version comparison to the detection'
        $message | Should -Not -Match 'Publish it as an Application'
    }

    It 'reads the 32-bit registry view when the detection is not marked 64-bit' {
        $detection = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Contoso\Tool'; ValueName = 'DisplayVersion'; DisplayVersion = '5.4.2' }
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })
        $rules.IsInstalled | Should -Match '^<bar:RegSzToVersion Key="HKEY_LOCAL_MACHINE" Subkey="SOFTWARE\\Contoso\\Tool" RegType32="true" '
        $rules.IsInstallable | Should -Be '<lar:And><bar:RegValueExists Key="HKEY_LOCAL_MACHINE" Subkey="SOFTWARE\Contoso\Tool" RegType32="true" Value="DisplayVersion" Type="REG_SZ" /><bar:RegSzToVersion Key="HKEY_LOCAL_MACHINE" Subkey="SOFTWARE\Contoso\Tool" RegType32="true" Value="DisplayVersion" Comparison="LessThan" Data="5.4.2.0" /></lar:And>'
    }

    It 'expands %ProgramFiles% the way a 32-bit process does unless the clause is marked 64-bit' -Skip:(-not ${env:ProgramFiles(x86)}) {
        $x86 = [pscustomobject]@{ Type = 'File'; FilePath = '%ProgramFiles%\Contoso'; FileName = 'tool.exe'; PropertyType = 'Version'; ExpectedValue = '5.4.2' }
        (ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $x86 })).IsInstalled |
            Should -Be '<bar:FileVersion Path="Contoso\tool.exe" Csidl="42" Comparison="GreaterThanOrEqualTo" Version="5.4.2.0" />'
        $x64 = [pscustomobject]@{ Type = 'File'; FilePath = '%ProgramFiles%\Contoso'; FileName = 'tool.exe'; PropertyType = 'Version'; ExpectedValue = '5.4.2'; Is64Bit = $true }
        (ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $x64 })).IsInstalled |
            Should -Be '<bar:FileVersion Path="Contoso\tool.exe" Csidl="38" Comparison="GreaterThanOrEqualTo" Version="5.4.2.0" />'
    }

    It 'maps a Program Files (x86) file version onto its CSIDL' {
        $detection = [pscustomobject]@{ Type = 'File'; FilePath = (Join-Path ${env:ProgramFiles(x86)} 'Contoso\Tool'); FileName = 'tool.exe'; PropertyType = 'Version'; Operator = 'GreaterEquals'; ExpectedValue = '5.4.2.17' }
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })
        $rules.IsInstalled | Should -Be '<bar:FileVersion Path="Contoso\Tool\tool.exe" Csidl="42" Comparison="GreaterThanOrEqualTo" Version="5.4.2.17" />'
        $rules.IsInstallable | Should -Be '<lar:And><bar:FileExists Path="Contoso\Tool\tool.exe" Csidl="42" /><bar:FileVersion Path="Contoso\Tool\tool.exe" Csidl="42" Comparison="LessThan" Version="5.4.2.17" /></lar:And>'
    }

    It 'keeps an absolute path outside the known folders' {
        $detection = [pscustomobject]@{ Type = 'File'; FilePath = 'D:\Tools\Contoso'; FileName = 'tool.exe'; PropertyType = 'Version'; ExpectedValue = '5.4.2' }
        (ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })).IsInstalled |
            Should -Be '<bar:FileVersion Path="D:\Tools\Contoso\tool.exe" Comparison="GreaterThanOrEqualTo" Version="5.4.2.0" />'
    }

    # A key existence clause joined by AND to a version clause. The values are
    # .NET host installer versions: three parts with a large last part.
    It 'offers a compound AND update only where the key exists and the version is older: <Name>' -TestCases @(
        @{ Name = 'neither'; Options = $false; HostVersion = $null; Installed = $false; Offered = $false }
        @{ Name = 'older version without the key'; Options = $false; HostVersion = '64.120.56788'; Installed = $false; Offered = $false }
        @{ Name = 'key without a version'; Options = $true; HostVersion = $null; Installed = $false; Offered = $false }
        @{ Name = 'key with an older version'; Options = $true; HostVersion = '64.120.56788'; Installed = $false; Offered = $true }
        @{ Name = 'key with the target version'; Options = $true; HostVersion = '64.124.58447'; Installed = $true; Offered = $false }
        @{ Name = 'key with a newer version'; Options = $true; HostVersion = '64.128.1000'; Installed = $true; Offered = $false }
    ) {
        param($Name, $Options, $HostVersion, $Installed, $Offered)
        $optionsKey = 'SOFTWARE\WOW6432Node\Microsoft\dotnet\host\options\8.0'
        $hostKey = 'SOFTWARE\Classes\Installer\Dependencies\Dotnet_CLI_SharedHost_8.0_x64'
        $detection = [pscustomobject]@{ Type = 'Compound'; Connector = 'And'; Clauses = @(
                [pscustomobject]@{ Type = 'RegistryKey'; RegistryKeyRelative = $optionsKey; Is64Bit = $true }
                [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $hostKey; ValueName = 'Version'; PropertyType = 'Version'; Operator = 'GreaterEquals'; ExpectedValue = '64.124.0'; Is64Bit = $true }) }
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })
        $registry = @{}
        if ($Options) { $registry[('{0}|64' -f $optionsKey)] = @{ OPT_NO_ANCM = '0' } }
        if ($HostVersion) { $registry[('{0}|64' -f $hostKey)] = @{ Version = $HostVersion } }
        $outcome = Get-TestOutcome -Rules $rules -Registry $registry
        $outcome.Installed | Should -Be $Installed
        $outcome.Offered | Should -Be $Offered
    }

    It 'reads a WSUS-only Add/Remove Programs entry through a registry loop in its view' {
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ WsusDetection = $script:HostingBundleArp })
        $rules.IsInstalled | Should -BeLike '<bar:RegKeyLoop Key="HKEY_LOCAL_MACHINE" Subkey="SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" RegType32="true" TrueIf="Any">*'
        $rules.IsInstalled | Should -BeLike '*<bar:RegSz Key="HKEY_LOOP_TARGET" Subkey="\" RegType32="true" Value="DisplayName" Comparison="BeginsWith" Data="Microsoft .NET 8.0." />*'
        $rules.IsInstalled | Should -BeLike '*Value="DisplayVersion" Comparison="GreaterThanOrEqualTo" Data="8.0.31.0" />*'
        $rules.IsInstallable | Should -BeLike '*<bar:RegValueExists Key="HKEY_LOOP_TARGET" Subkey="\" RegType32="true" Value="DisplayVersion" Type="REG_SZ" />*'
        Test-WellFormedRule -Rule $rules.IsInstalled | Should -BeTrue
        Test-WellFormedRule -Rule $rules.IsInstallable | Should -BeTrue
    }

    It 'offers the hosting bundle update only over an older hosting bundle entry: <Name>' -TestCases @(
        @{ Name = 'no .NET entry'; Entries = @(); Installed = $false; Offered = $false }
        @{ Name = 'older bundle'; Entries = @(@{ Key = '{da035e95-ecb4-4cec-854c-87f6b29fde6b}'; Name = 'Microsoft .NET 8.0.30 - Windows Server Hosting'; Version = '8.0.30.26373' }); Installed = $false; Offered = $true }
        @{ Name = 'target bundle'; Entries = @(@{ Key = '{b3c4f892-1e5d-4df6-ab9c-7bc005bdee67}'; Name = 'Microsoft .NET 8.0.31 - Windows Server Hosting'; Version = '8.0.31.26421' }); Installed = $true; Offered = $false }
        @{ Name = 'newer bundle'; Entries = @(@{ Key = '{0c1d2e3f-0000-4000-8000-000000000001}'; Name = 'Microsoft .NET 8.0.32 - Windows Server Hosting'; Version = '8.0.32.26500' }); Installed = $true; Offered = $false }
        @{ Name = 'older bundle beside a newer Desktop Runtime'; Entries = @(
                @{ Key = '{da035e95-ecb4-4cec-854c-87f6b29fde6b}'; Name = 'Microsoft .NET 8.0.30 - Windows Server Hosting'; Version = '8.0.30.26373' }
                @{ Key = '{11111111-2222-4333-8444-555555555555}'; Name = 'Microsoft Windows Desktop Runtime - 8.0.31 (x64)'; Version = '8.0.31.26421' }); Installed = $false; Offered = $true }
        @{ Name = 'Desktop Runtime only'; Entries = @(@{ Key = '{11111111-2222-4333-8444-555555555555}'; Name = 'Microsoft Windows Desktop Runtime - 8.0.30 (x64)'; Version = '8.0.30.26373' }); Installed = $false; Offered = $false }
        @{ Name = '.NET 10 bundle only'; Entries = @(@{ Key = '{20FB5C72-3DE9-415A-B800-494BC18BEC76}'; Name = 'Microsoft .NET 10.0.11 - Windows Server Hosting'; Version = '10.0.11.26373' }); Installed = $false; Offered = $false }
        @{ Name = 'bundle entry in the 64-bit view'; Entries = @(@{ Key = '{da035e95-ecb4-4cec-854c-87f6b29fde6b}'; Name = 'Microsoft .NET 8.0.30 - Windows Server Hosting'; Version = '8.0.30.26373'; View = '64' }); Installed = $false; Offered = $false }
        @{ Name = 'bundle entry without a version'; Entries = @(@{ Key = '{da035e95-ecb4-4cec-854c-87f6b29fde6b}'; Name = 'Microsoft .NET 8.0.30 - Windows Server Hosting' }); Installed = $false; Offered = $false }
        @{ Name = 'another publisher'; Entries = @(@{ Key = '{da035e95-ecb4-4cec-854c-87f6b29fde6b}'; Name = 'Microsoft .NET 8.0.30 - Windows Server Hosting'; Version = '8.0.30.26373'; Publisher = 'Contoso' }); Installed = $false; Offered = $false }
        @{ Name = 'stale older entry beside the current one'; Entries = @(
                @{ Key = '{da035e95-ecb4-4cec-854c-87f6b29fde6b}'; Name = 'Microsoft .NET 8.0.30 - Windows Server Hosting'; Version = '8.0.30.26373' }
                @{ Key = '{b3c4f892-1e5d-4df6-ab9c-7bc005bdee67}'; Name = 'Microsoft .NET 8.0.31 - Windows Server Hosting'; Version = '8.0.31.26421' }); Installed = $true; Offered = $false }
    ) {
        param($Name, $Entries, $Installed, $Offered)
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ WsusDetection = $script:HostingBundleArp })
        $outcome = Get-TestOutcome -Rules $rules -Registry (New-TestArpRegistry -Entries $Entries)
        $outcome.Installed | Should -Be $Installed
        $outcome.Offered | Should -Be $Offered
    }

    It 'reads both registry views for view Both: <View>' -TestCases @(@{ View = '32' }, @{ View = '64' }) {
        param($View)
        $detection = [pscustomobject]@{ Type = 'ArpEntry'; View = 'Both'; DisplayNamePrefix = 'Contoso Tool '; Version = '5.4.2' }
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ WsusDetection = $detection })
        (Get-TestOutcome -Rules $rules -Registry (New-TestArpRegistry -Entries @(@{ Key = 'ContosoTool'; Name = 'Contoso Tool 5.4.1'; Version = '5.4.1'; Publisher = 'Contoso'; View = $View }))).Offered | Should -BeTrue
        (Get-TestOutcome -Rules $rules -Registry (New-TestArpRegistry -Entries @(@{ Key = 'ContosoTool'; Name = 'Contoso Tool 5.4.2'; Version = '5.4.2'; Publisher = 'Contoso'; View = $View }))).Installed | Should -BeTrue
    }

    It 'publishes a manifest whose own detection WSUS cannot map when it names an Add/Remove Programs entry' {
        $manifest = New-TestManifest -Overrides @{
            Detection     = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:ProductKey; DisplayVersion = '5.4.2'; Is64Bit = $true }
            WsusDetection = [pscustomobject]@{ Type = 'ArpEntry'; View = '64'; DisplayNamePrefix = 'Contoso Tool'; Publisher = 'Contoso'; Version = '5.4.2' }
        }
        { ConvertTo-WsusApplicabilityRules -Manifest $manifest } | Should -Not -Throw
        @(Get-WsusCompatibilityFindings -Manifest $manifest -ContentFolder $script:Content | Where-Object Code -eq 'DetectionNotMappable').Count | Should -Be 0
    }

    It 'reads the Add/Remove Programs entry that the staged MSI registers: <Name>' -TestCases @(
        @{ Name = 'version inside the name'; Msi = @{ ProductName = '7-Zip 24.08 (x64 edition)'; ProductVersion = '24.08.00.0'; Manufacturer = 'Igor Pavlov'; Platform = 'x64' }
            Expect = @{ View = '64'; DisplayNamePrefix = '7-Zip '; DisplayNameSuffix = ' (x64 edition)'; DisplayName = ''; Version = '24.08.00.0'; Publisher = 'Igor Pavlov'; WindowsInstaller = 'True' } }
        @{ Name = 'version at the end of the name'; Msi = @{ ProductName = 'Contoso Tool 5.4.2'; ProductVersion = '5.4.2.0'; Manufacturer = 'Contoso'; Platform = 'Intel' }
            Expect = @{ View = '32'; DisplayNamePrefix = 'Contoso Tool '; DisplayNameSuffix = ''; DisplayName = ''; Version = '5.4.2.0'; Publisher = 'Contoso'; WindowsInstaller = 'True' } }
        @{ Name = 'no version in the name'; Msi = @{ ProductName = 'Google Chrome'; ProductVersion = '131.0.6778.86'; Manufacturer = 'Google LLC'; Platform = 'x64' }
            Expect = @{ View = '64'; DisplayNamePrefix = ''; DisplayNameSuffix = ''; DisplayName = 'Google Chrome'; Version = '131.0.6778.86'; Publisher = 'Google LLC'; WindowsInstaller = 'True' } }
    ) {
        param($Name, $Msi, $Expect)
        $script:MsiInfo = [pscustomobject]$Msi
        Mock -ModuleName AppPackagerWsus Get-WsusMsiInfo { $script:MsiInfo }
        $manifest = New-TestManifest -Overrides @{
            InstallerType = 'MSI'; InstallerFile = 'product.msi'
            Detection     = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:ProductKey; DisplayVersion = '5.4.2'; Is64Bit = $true }
        }
        $rules = ConvertTo-WsusApplicabilityRules -Manifest $manifest -ContentFolder $script:Content
        $rules.Source | Should -Be 'MsiEntry'
        foreach ($key in $Expect.Keys) { [string]$rules.Entry.PSObject.Properties[$key].Value | Should -Be $Expect[$key] -Because $key }
        Test-WellFormedRule -Rule $rules.IsInstallable | Should -BeTrue
    }

    It 'offers an MSI product update only over an older entry of that product: <Name>' -TestCases @(
        @{ Name = 'older version'; Entry = @{ Key = '{23170F69-40C1-2702-2407-000001000000}'; Name = '7-Zip 24.07 (x64 edition)'; Version = '24.07.00.0'; Publisher = 'Igor Pavlov'; View = '64'; WindowsInstaller = 1 }; Installed = $false; Offered = $true }
        @{ Name = 'target version'; Entry = @{ Key = '{23170F69-40C1-2702-2408-000001000000}'; Name = '7-Zip 24.08 (x64 edition)'; Version = '24.08.00.0'; Publisher = 'Igor Pavlov'; View = '64'; WindowsInstaller = 1 }; Installed = $true; Offered = $false }
        @{ Name = 'the 32-bit edition'; Entry = @{ Key = '{23170F69-40C1-2701-2407-000001000000}'; Name = '7-Zip 24.07'; Version = '24.07.00.0'; Publisher = 'Igor Pavlov'; View = '32'; WindowsInstaller = 1 }; Installed = $false; Offered = $false }
        @{ Name = 'another product'; Entry = @{ Key = '{11111111-2222-4333-8444-555555555555}'; Name = '7-Zip Helper 1.0 (x64 edition)'; Version = '1.0'; Publisher = 'Contoso'; View = '64'; WindowsInstaller = 1 }; Installed = $false; Offered = $false }
        @{ Name = 'an EXE-installed copy with the same name'; Entry = @{ Key = '7-Zip'; Name = '7-Zip 24.07 (x64 edition)'; Version = '24.07.00.0'; Publisher = 'Igor Pavlov'; View = '64' }; Installed = $false; Offered = $false }
    ) {
        param($Name, $Entry, $Installed, $Offered)
        $script:MsiInfo = [pscustomobject]@{ ProductName = '7-Zip 24.08 (x64 edition)'; ProductVersion = '24.08.00.0'; Manufacturer = 'Igor Pavlov'; Platform = 'x64' }
        Mock -ModuleName AppPackagerWsus Get-WsusMsiInfo { $script:MsiInfo }
        $manifest = New-TestManifest -Overrides @{
            InstallerType = 'MSI'; InstallerFile = 'product.msi'
            Detection     = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:ProductKey; DisplayVersion = '24.08.00.0'; Is64Bit = $true }
        }
        $outcome = Get-TestOutcome -Rules (ConvertTo-WsusApplicabilityRules -Manifest $manifest -ContentFolder $script:Content) -Registry (New-TestArpRegistry -Entries @($Entry))
        $outcome.Installed | Should -Be $Installed
        $outcome.Offered | Should -Be $Offered
    }

    It 'never offers an MSI update to an EXE installation whose name starts the same way' {
        # KeePass: the MSI registers "KeePass <version>" and the EXE setup
        # registers "KeePass Password Safe <version>", both by the same vendor.
        $script:MsiInfo = [pscustomobject]@{ ProductName = 'KeePass 2.59'; ProductVersion = '2.59.0.0'; Manufacturer = 'Dominik Reichl'; Platform = 'Intel' }
        Mock -ModuleName AppPackagerWsus Get-WsusMsiInfo { $script:MsiInfo }
        $manifest = New-TestManifest -Overrides @{
            InstallerType = 'MSI'; InstallerFile = 'product.msi'
            Detection     = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:ProductKey; DisplayVersion = '2.59.0.0'; Is64Bit = $false }
        }
        $rules = ConvertTo-WsusApplicabilityRules -Manifest $manifest -ContentFolder $script:Content
        $rules.Entry.WindowsInstaller | Should -BeTrue
        $exe = @{ Key = 'KeePassPasswordSafe2_is1'; Name = 'KeePass Password Safe 2.58'; Version = '2.58'; Publisher = 'Dominik Reichl'; View = '32' }
        $msi = @{ Key = '{E1E1A3B6-ED2C-4D53-B5A8-2F6F0B2E0AB3}'; Name = 'KeePass 2.58'; Version = '2.58.0.0'; Publisher = 'Dominik Reichl'; View = '32'; WindowsInstaller = 1 }
        (Get-TestOutcome -Rules $rules -Registry (New-TestArpRegistry -Entries @($exe))).Offered | Should -BeFalse
        (Get-TestOutcome -Rules $rules -Registry (New-TestArpRegistry -Entries @($msi))).Offered | Should -BeTrue
        (Get-TestOutcome -Rules $rules -Registry (New-TestArpRegistry -Entries @($exe, $msi))).Offered | Should -BeTrue
    }

    It 'requires the Windows Installer flag when a WsusDetection block asks for it' {
        $detection = [pscustomobject]@{ Type = 'ArpEntry'; View = '32'; DisplayNamePrefix = 'Contoso Tool '; Publisher = 'Contoso'; WindowsInstaller = $true; Version = '5.4.2' }
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ WsusDetection = $detection })
        $rules.IsInstalled | Should -BeLike '*<bar:RegDword Key="HKEY_LOOP_TARGET" Subkey="\" RegType32="true" Value="WindowsInstaller" Comparison="EqualTo" Data="1" />*'
        $older = @{ Key = 'ContosoTool'; Name = 'Contoso Tool 5.4.1'; Version = '5.4.1'; Publisher = 'Contoso'; View = '32'; WindowsInstaller = 1 }
        (Get-TestOutcome -Rules $rules -Registry (New-TestArpRegistry -Entries @($older))).Offered | Should -BeTrue
        $older.Remove('WindowsInstaller')
        (Get-TestOutcome -Rules $rules -Registry (New-TestArpRegistry -Entries @($older))).Offered | Should -BeFalse
    }

    It 'keeps a detection that WSUS can map, even when the installer is an MSI' {
        Mock -ModuleName AppPackagerWsus Get-WsusMsiInfo { throw 'must not read the MSI' }
        $manifest = New-TestManifest -Overrides @{ InstallerType = 'MSI'; InstallerFile = 'product.msi' }
        (ConvertTo-WsusApplicabilityRules -Manifest $manifest -ContentFolder $script:Content).Source | Should -Be 'Detection'
    }

    It 'still refuses a detection WSUS cannot map when the installer is not an MSI or no content folder is given' {
        $blocked = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:ProductKey; DisplayVersion = '5.4.2'; Is64Bit = $true }
        { ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $blocked }) -ContentFolder $script:Content } | Should -Throw '*Windows Installer product key*'
        { ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ InstallerType = 'MSI'; InstallerFile = 'product.msi'; Detection = $blocked }) } | Should -Throw '*Windows Installer product key*'
    }

    It 'reports the MSI entry as an Info finding instead of a blocking one' {
        $script:MsiInfo = [pscustomobject]@{ ProductName = 'Contoso Tool 5.4.2'; ProductVersion = '5.4.2.0'; Manufacturer = 'Contoso'; Platform = 'x64' }
        Mock -ModuleName AppPackagerWsus Get-WsusMsiInfo { $script:MsiInfo }
        $manifest = New-TestManifest -Overrides @{
            InstallerType = 'MSI'; InstallerFile = 'product.msi'
            Detection     = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = $script:ProductKey; DisplayVersion = '5.4.2'; Is64Bit = $true }
        }
        $findings = @(Get-WsusCompatibilityFindings -Manifest $manifest -ContentFolder $script:Content)
        @($findings | Where-Object Code -eq 'DetectionNotMappable').Count | Should -Be 0
        $info = @($findings | Where-Object Code -eq 'WsusRulesFromMsi')
        $info.Count | Should -Be 1
        $info[0].Severity | Should -Be 'Info'
        $info[0].Message | Should -Match "'Contoso Tool \*' by Contoso, 64-bit view"
    }

    It 'reads ProductName, ProductVersion, Manufacturer and the platform from an MSI file' {
        # A PSObject-wrapped path reaches the COM call as a type mismatch.
        $msi = [string](Join-Path $TestDrive 'probe.msi')
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $installer, @($msi, 3))
        $statements = @(
            'CREATE TABLE `Property` (`Property` CHAR(72) NOT NULL, `Value` LONGCHAR NOT NULL LOCALIZABLE PRIMARY KEY `Property`)'
            "INSERT INTO ``Property`` (``Property``, ``Value``) VALUES ('ProductName', 'Contoso Tool 5.4.2')"
            "INSERT INTO ``Property`` (``Property``, ``Value``) VALUES ('ProductVersion', '5.4.2.0')"
            "INSERT INTO ``Property`` (``Property``, ``Value``) VALUES ('Manufacturer', 'Contoso')"
        )
        foreach ($sql in $statements) {
            $view = $db.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $db, @($sql))
            [void]$view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
            [void]$view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null)
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($view)
        }
        $summary = $db.GetType().InvokeMember('SummaryInformation', 'GetProperty', $null, $db, 4)
        [void]$summary.GetType().InvokeMember('Property', 'SetProperty', $null, $summary, @(7, 'x64;1033'))
        [void]$summary.GetType().InvokeMember('Persist', 'InvokeMethod', $null, $summary, $null)
        [void]$db.GetType().InvokeMember('Commit', 'InvokeMethod', $null, $db, $null)
        foreach ($o in @($summary, $db, $installer)) { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($o) }
        $info = InModuleScope AppPackagerWsus -Parameters @{ Path = $msi } { param($Path) Get-WsusMsiInfo -Path $Path }
        $info.ProductName | Should -Be 'Contoso Tool 5.4.2'
        $info.ProductVersion | Should -Be '5.4.2.0'
        $info.Manufacturer | Should -Be 'Contoso'
        $info.Platform | Should -Be 'x64'
        { [System.IO.File]::Delete($msi) } | Should -Not -Throw
    }

    It 'refuses a WsusDetection it cannot map: <Name>' -TestCases @(
        @{ Name = 'unknown type'; Detection = [pscustomobject]@{ Type = 'File'; View = '32'; DisplayNamePrefix = 'X'; Version = '1.0' }; Message = '*type ArpEntry*' }
        @{ Name = 'no DisplayName prefix'; Detection = [pscustomobject]@{ Type = 'ArpEntry'; View = '32'; Version = '1.0' }; Message = '*DisplayNamePrefix*' }
        @{ Name = 'unknown view'; Detection = [pscustomobject]@{ Type = 'ArpEntry'; View = 'x86'; DisplayNamePrefix = 'X'; Version = '1.0' }; Message = '*not 32, 64 or Both*' }
        @{ Name = 'version that is not numeric'; Detection = [pscustomobject]@{ Type = 'ArpEntry'; View = '32'; DisplayNamePrefix = 'X'; Version = 'latest' }; Message = '*not a version*' }
    ) {
        param($Name, $Detection, $Message)
        $manifest = New-TestManifest -Overrides @{ WsusDetection = $Detection }
        { ConvertTo-WsusApplicabilityRules -Manifest $manifest } | Should -Throw $Message
        @(Get-WsusCompatibilityFindings -Manifest $manifest -ContentFolder $script:Content | Where-Object { $_.Code -eq 'DetectionNotMappable' -and $_.Severity -eq 'Blocking' }).Count | Should -Be 1
    }

    It 'maps the Visual C++ runtime file in <Folder> onto CSIDL <Csidl>' -TestCases @(
        @{ Folder = '%SystemRoot%\System32'; Csidl = '37'; Path = 'vcruntime140.dll' }
        @{ Folder = '%SystemRoot%\SysWOW64'; Csidl = '36'; Path = 'SysWOW64\vcruntime140.dll' }
    ) {
        param($Folder, $Csidl, $Path)
        $detection = [pscustomobject]@{ Type = 'File'; FilePath = $Folder; FileName = 'vcruntime140.dll'; PropertyType = 'Version'; Operator = 'GreaterEquals'; ExpectedValue = '14.51.36247.0'; Is64Bit = $true }
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })
        $rules.IsInstalled | Should -Be ('<bar:FileVersion Path="{0}" Csidl="{1}" Comparison="GreaterThanOrEqualTo" Version="14.51.36247.0" />' -f $Path, $Csidl)
        $rules.IsInstallable | Should -Be ('<lar:And><bar:FileExists Path="{0}" Csidl="{1}" /><bar:FileVersion Path="{0}" Csidl="{1}" Comparison="LessThan" Version="14.51.36247.0" /></lar:And>' -f $Path, $Csidl)
    }

    It 'refuses detections WSUS cannot evaluate' -TestCases @(
        @{ Name = 'per-user hive'; Detection = [pscustomobject]@{ Type = 'RegistryKey'; Hive = 'CurrentUser'; RegistryKeyRelative = 'Software\Contoso' }; Message = '*HKEY_CURRENT_USER*' }
        @{ Name = 'script'; Detection = [pscustomobject]@{ Type = 'Script'; ScriptText = 'exit 0' }; Message = '*Script detection*' }
        @{ Name = 'per-user path'; Detection = [pscustomobject]@{ Type = 'File'; FilePath = '%LOCALAPPDATA%\Contoso'; FileName = 'tool.exe'; PropertyType = 'Version'; ExpectedValue = '1.0' }; Message = '*per-user path*' }
        @{ Name = 'grouped compound'; Detection = [pscustomobject]@{ Type = 'Compound'; Connector = 'And'; GroupSizes = @(1, 1); Clauses = @() }; Message = '*Grouped compound*' }
        @{ Name = 'less-than operator'; Detection = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\C'; ExpectedValue = '1.0'; Operator = 'LessThan'; Is64Bit = $true }; Message = "*operator 'LessThan'*" }
        @{ Name = 'no detection'; Detection = $null; Message = '*carries no detection*' }
    ) {
        param($Name, $Detection, $Message)
        { ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $Detection }) } | Should -Throw $Message
    }

    It 'escapes XML metacharacters in keys, values and paths' {
        $detection = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\A&B <"x">'; ValueName = "V'1"; ExpectedValue = '1.0'; Operator = 'IsEquals'; Is64Bit = $true }
        $rules = ConvertTo-WsusApplicabilityRules -Manifest (New-TestManifest -Overrides @{ Detection = $detection })
        Test-WellFormedRule -Rule $rules.IsInstalled | Should -BeTrue
        Test-WellFormedRule -Rule $rules.IsInstallable | Should -BeTrue
        $rules.IsInstalled | Should -Match 'Subkey="SOFTWARE\\A&amp;B &lt;&quot;x&quot;&gt;"'
        $rules.IsInstallable | Should -Match 'Value="V&apos;1"'
    }
}

Describe 'Get-WsusCompatibilityFindings' {
    It 'reports no blocking finding for a plain silent EXE' {
        $findings = @(Get-WsusCompatibilityFindings -Manifest (New-TestManifest) -ContentFolder $script:Content)
        @($findings | Where-Object Severity -eq 'Blocking').Count | Should -Be 0
        @($findings | Where-Object Code -eq 'UninstallNotPublished').Count | Should -Be 1
    }

    It 'blocks <Code>' -TestCases @(
        @{ Code = 'MultipleDeploymentTypes'; Overrides = @{ DeploymentTypes = @([pscustomobject]@{ NameSuffix = 'x64' }) } }
        @{ Code = 'InstallerTypeUnsupported'; Overrides = @{ InstallerType = 'MSIX' } }
        @{ Code = 'PerUserInstall'; Overrides = @{ InstallationBehaviorType = 'InstallForUser' } }
        @{ Code = 'InteractiveInstall'; Overrides = @{ RequireUserInteraction = $true } }
        @{ Code = 'CustomInstall'; Overrides = @{ InstallCommandLine = 'Deploy-Application.exe -DeploymentType Install' } }
        @{ Code = 'CustomInstall'; Overrides = @{ CustomAssets = @([pscustomobject]@{ Category = 'InstallBefore'; RelativePath = 'install-before.ps1' }) } }
        @{ Code = 'CustomInstall'; Overrides = @{ CommandOverrides = [pscustomobject]@{ Install = 'setup.exe /quiet'; Uninstall = '' } } }
        @{ Code = 'InstallerMissing'; Overrides = @{ InstallerFile = 'absent.exe' } }
        @{ Code = 'DetectionNotMappable'; Overrides = @{ Detection = [pscustomobject]@{ Type = 'Script'; ScriptText = 'exit 0' } } }
    ) {
        param($Code, $Overrides)
        $findings = @(Get-WsusCompatibilityFindings -Manifest (New-TestManifest -Overrides $Overrides) -ContentFolder $script:Content)
        @($findings | Where-Object { $_.Severity -eq 'Blocking' -and $_.Code -eq $Code }).Count | Should -Be 1
    }

    It 'does not block an uninstall-only command override or the default install entry' {
        $manifest = New-TestManifest -Overrides @{ InstallCommandLine = 'install.bat'; CommandOverrides = [pscustomobject]@{ Install = ''; Uninstall = 'uninstall.exe /S' } }
        @(Get-WsusCompatibilityFindings -Manifest $manifest -ContentFolder $script:Content | Where-Object Severity -eq 'Blocking').Count | Should -Be 0
    }

    It 'asks for review of requirement rules and running processes' {
        $manifest = New-TestManifest -Overrides @{ Requirements = @([pscustomobject]@{ RuleId = 'cpu-arch' }); RunningProcess = @('tool') }
        $codes = @(Get-WsusCompatibilityFindings -Manifest $manifest -ContentFolder $script:Content | Where-Object Severity -eq 'Review' | ForEach-Object Code)
        $codes | Should -Contain 'RequirementsNotTranslated'
        $codes | Should -Contain 'RunningProcessNotClosed'
    }

    It 'blocks an install.ps1 that does more than run the installer' {
        $folder = Join-Path $script:Root ('custom-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Copy-Item -Path (Join-Path $script:Content '*') -Destination $folder
        [System.IO.File]::WriteAllText((Join-Path $folder 'install.ps1'), ($script:StandardInstallScript + "`r`n[Environment]::SetEnvironmentVariable('DBEAVER_AI_DISABLED', 'true', 'Machine')"))
        $finding = @(Get-WsusCompatibilityFindings -Manifest (New-TestManifest) -ContentFolder $folder | Where-Object { $_.Code -eq 'CustomInstall' })
        $finding.Count | Should -Be 1
        $finding[0].Severity | Should -Be 'Blocking'
        $finding[0].Message | Should -Match 'install\.ps1 does more than run the installer: \[Environment\]::SetEnvironmentVariable'
    }

    It 'blocks recorded payload files in a subfolder or missing from the stage' {
        $folder = Join-Path $script:Root ('nested-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path (Join-Path $folder 'Data') -Force | Out-Null
        Copy-Item -Path (Join-Path $script:Content '*') -Destination $folder
        [System.IO.File]::WriteAllText((Join-Path $folder 'Data\x.cab'), 'cab')
        $hashes = @('setup.exe', 'Data\x.cab', 'gone.cab') | ForEach-Object { [pscustomobject]@{ RelativePath = $_; Sha256 = ('0' * 64); Size = 1 } }
        $findings = @(Get-WsusCompatibilityFindings -Manifest (New-TestManifest -Overrides @{ FileHashes = $hashes }) -ContentFolder $folder)
        ($findings | Where-Object Code -eq 'PayloadInSubfolder').Message | Should -Match 'Data\\x\.cab'
        ($findings | Where-Object Code -eq 'PayloadMissing').Message | Should -Match 'gone\.cab'
    }

    It 'blocks a detection whose expected value is not a numeric version' {
        $detection = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Contoso'; ValueName = 'DisplayVersion'; ExpectedValue = '2026.08.2-200'; Operator = 'IsEquals'; Is64Bit = $true }
        $finding = @(Get-WsusCompatibilityFindings -Manifest (New-TestManifest -Overrides @{ Detection = $detection }) -ContentFolder $script:Content | Where-Object Code -eq 'DetectionNotMappable')
        $finding.Count | Should -Be 1
        $finding[0].Severity | Should -Be 'Blocking'
        $finding[0].Message | Should -Match '2026\.08\.2-200'
    }
}

Describe 'Publish-WsusSoftwareUpdate' {
    BeforeAll {
        function New-HashedManifest {
            param([hashtable]$Overrides = @{})
            $manifest = New-TestManifest -Overrides $Overrides
            $hashes = foreach ($file in @(Get-ChildItem -LiteralPath $script:Content -File | Where-Object Name -ne 'stage-manifest.json')) {
                [pscustomobject]@{ RelativePath = $file.Name; Sha256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash; Size = $file.Length }
            }
            $manifest | Add-Member -NotePropertyName FileHashes -NotePropertyValue @($hashes) -Force
            $manifest | Add-Member -NotePropertyName BuildId -NotePropertyValue '20260929-101500-abcdef01' -Force
            return $manifest
        }

        function New-FakeUpdate {
            param([guid]$Id, [string]$Title, [string]$Description, [bool]$Declined = $false)
            [pscustomobject]@{
                Id           = [pscustomobject]@{ UpdateId = $Id }
                Title        = $Title
                Description  = $Description
                IsDeclined   = $Declined
                IsSuperseded = $false
                CreationDate = [datetime]'2026-09-01'
            }
        }

        $script:Settings = @{ ServerName = 'wsus01.contoso.com'; PortNumber = 8531; UseSsl = $true; PackageType = 'Update'; Classification = 'Updates' }
        $script:Identity = 'AppPackager:catalog:package-contosotool/default'
    }

    BeforeEach {
        $script:Plans = New-Object System.Collections.Generic.List[object]
        $script:Published = New-Object System.Collections.Generic.List[object]
        $script:Approvals = New-Object System.Collections.Generic.List[object]
        $script:Declines = New-Object System.Collections.Generic.List[object]
        $script:ServerUpdates = @()
        $script:ExistingIds = @()

        Mock -ModuleName AppPackagerWsus Get-WsusServerConnection { [pscustomobject]@{ Name = 'wsus01.contoso.com'; Version = '10.0.20348.1' } }
        Mock -ModuleName AppPackagerWsus Get-WsusUserRoleName { 'Administrator' }
        Mock -ModuleName AppPackagerWsus Get-WsusLocalPublishingCabLimitMegabytes { 384 }
        Mock -ModuleName AppPackagerWsus Get-WsusSigningCertificateObject {
            [pscustomobject]@{
                Subject = 'CN=WSUS Publishers Self-signed'; Issuer = 'CN=WSUS Publishers Self-signed'; Thumbprint = 'AB12'
                NotBefore = (Get-Date).AddDays(-1); NotAfter = (Get-Date).AddYears(1)
                PublicKey = [pscustomobject]@{ Key = [pscustomobject]@{ KeySize = 2048 } }
            }
        }
        Mock -ModuleName AppPackagerWsus Get-WsusLocallyPublishedUpdateObjects { $script:ServerUpdates }
        Mock -ModuleName AppPackagerWsus Get-WsusUpdateObject {
            if ($script:ExistingIds -contains $PackageId) { return (New-FakeUpdate -Id $PackageId -Title 'existing' -Description '') }
            return $null
        }
        Mock -ModuleName AppPackagerWsus New-WsusSoftwareDistributionPackageFile {
            $script:Plans.Add([pscustomobject]@{ Plan = $Plan; Files = @(Get-ChildItem -LiteralPath $Plan.SourceFolder -File | ForEach-Object Name) })
            [System.IO.File]::WriteAllText($Path, '<sdp/>')
        }
        Mock -ModuleName AppPackagerWsus Invoke-WsusPublisher {
            $script:Published.Add([pscustomobject]@{ SdpPath = $SdpPath; SourceFolder = $SourceFolder })
            $script:ExistingIds += [guid]$script:Plans[-1].Plan.PackageId
        }
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateApproval { $script:Approvals.Add($GroupName) }
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateDecline { $script:Declines.Add([string]$Update.Id.UpdateId) }
    }

    It 'refuses a payload that changes between the first integrity check and the package copy' {
        $manifest = New-HashedManifest
        $installer = Join-Path $script:Content ([string]$manifest.InstallerFile)
        $original = [System.IO.File]::ReadAllBytes($installer)
        $script:ChangeTarget = $installer
        Mock -ModuleName AppPackagerWsus Get-WsusLocallyPublishedUpdateObjects {
            [System.IO.File]::AppendAllText($script:ChangeTarget, 'changed')
            @()
        }
        try {
            { Publish-WsusSoftwareUpdate -Manifest $manifest -ContentFolder $script:Content -Settings $script:Settings } | Should -Throw '*changed after it was staged*'
            $script:Published.Count | Should -Be 0
        }
        finally { [System.IO.File]::WriteAllBytes($installer, $original) }
    }

    It 'names database maintenance for a SQL command timeout' {
        InModuleScope AppPackagerWsus {
            Get-WsusPublishFailureHint -Message 'Execution Timeout Expired.  The timeout period elapsed prior to completion of the operation or the server is not responding.' |
                Should -Match 'WSUS database maintenance'
            Get-WsusPublishFailureHint -Message 'Access is denied.' | Should -BeNullOrEmpty
        }
    }

    It 'publishes a new update built from the manifest' {
        $result = Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $script:Settings
        $expectedId = [guid]$script:Plans[0].Plan.PackageId
        $result.Outcome | Should -Be 'Published'
        $result.PackageId | Should -Be $expectedId
        $expectedId | Should -Not -Be ([guid]::Empty)
        $plan = $script:Plans[0].Plan
        $plan.Title | Should -Be 'Contoso Tool 5.4 5.4.2'
        $plan.ProductName | Should -Be 'AppPackager Applications'
        $plan.VendorName | Should -Be 'AppPackager'
        $plan.CommandLine | Should -Be '/S'
        $plan.InstallerFile | Should -Be 'setup.exe'
        $plan.Description | Should -Match ([regex]::Escape((Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update)))
        $plan.Description | Should -Match 'build 20260929-101500-abcdef01'
        $plan.IsInstalled | Should -Match '^<bar:RegSzToVersion .*Comparison="GreaterThanOrEqualTo"'
        $plan.IsInstallable | Should -Match '^<lar:And><bar:RegValueExists .*<bar:RegSzToVersion .*Comparison="LessThan"'
        $plan.PSObject.Properties['PackageType'] | Should -BeNullOrEmpty
        $plan.Description | Should -Match '; type Update; version 5\.4\.2\r?$'
        @($script:Plans[0].Files | Sort-Object) | Should -Be @('custom.mst', 'product.msi', 'setup.exe')
        Should -Invoke -ModuleName AppPackagerWsus Invoke-WsusPublisher -Times 1 -Exactly -ParameterFilter { $PackageId -eq $expectedId }
        Test-Path -LiteralPath (Split-Path -Parent $script:Published[0].SourceFolder) | Should -BeFalse
        $result.Message | Should -Not -Match 'rules read'
    }

    It 'names the Add/Remove Programs entry in the result when the rules come from the staged MSI' {
        $script:MsiInfo = [pscustomobject]@{ ProductName = 'Contoso Tool 5.4.2'; ProductVersion = '5.4.2.0'; Manufacturer = 'Contoso'; Platform = 'x64' }
        Mock -ModuleName AppPackagerWsus Get-WsusMsiInfo { $script:MsiInfo }
        $manifest = New-HashedManifest -Overrides @{
            InstallerType = 'MSI'; InstallerFile = 'product.msi'
            Detection     = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{23170F69-40C1-2702-2408-000001000000}'; DisplayVersion = '5.4.2'; Is64Bit = $true }
        }
        $result = Publish-WsusSoftwareUpdate -Manifest $manifest -ContentFolder $script:Content -Settings $script:Settings
        $result.Outcome | Should -Be 'Published'
        $result.Message | Should -Match ([regex]::Escape("; rules read the Add/Remove Programs entry that the MSI registers: 'Contoso Tool *' by Contoso, 64-bit view"))
    }

    It 'supersedes only lower versions of the same identity and type and declines them when asked' {
        $older = [guid]'00000000-0000-0000-0000-000000000541'
        $declinedOlder = [guid]'00000000-0000-0000-0000-000000000540'
        $script:ServerUpdates = @(
            (New-FakeUpdate -Id $older -Title 'Contoso Tool 5.4.1' -Description ("x`r`n" + (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update -Version '5.4.1')))
            (New-FakeUpdate -Id $declinedOlder -Title 'Contoso Tool 5.4.0' -Description (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update -Version '5.4.0') -Declined $true)
            (New-FakeUpdate -Id ([guid]::NewGuid()) -Title 'Contoso Tool 5.4.3' -Description (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update -Version '5.4.3'))
            (New-FakeUpdate -Id ([guid]::NewGuid()) -Title 'Contoso Tool unversioned' -Description (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update))
            ((New-FakeUpdate -Id ([guid]::NewGuid()) -Title 'Contoso Tool 5.3' -Description (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update -Version '5.3') -Declined $true) |
                Add-Member -NotePropertyName PublicationState -NotePropertyValue 'Expired' -PassThru)
            (New-FakeUpdate -Id ([guid]::NewGuid()) -Title 'Contoso Tool app' -Description (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Application -Version '5.4.1'))
            (New-FakeUpdate -Id ([guid]::NewGuid()) -Title 'Other' -Description (Get-WsusIdentityLine -IdentityTag 'AppPackager:catalog:package-other/default' -PackageType Update -Version '1.0'))
            (New-FakeUpdate -Id ([guid]::NewGuid()) -Title 'Foreign' -Description 'Published elsewhere')
        )
        $settings = @{} + $script:Settings
        $settings.DeclineSuperseded = $true
        $result = Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $settings
        @($script:Plans[0].Plan.SupersededPackageIds) | Should -Be @($older, $declinedOlder)
        @($result.Superseded).Count | Should -Be 2
        $script:Declines.ToArray() | Should -Be @([string]$older)
        $result.Message | Should -Match 'declined 1 earlier version'
        $result.Message | Should -Match '1 update\(s\) of the retired type Application stay on the server'
        @($result.Warnings).Count | Should -Be 0
    }

    It 'reports an existing update without publishing again and still applies the approval and the decline' {
        $current = [guid]'00000000-0000-0000-0000-000000000542'
        $older = [guid]'00000000-0000-0000-0000-000000000541'
        $script:ServerUpdates = @(
            (New-FakeUpdate -Id $current -Title 'Contoso Tool 5.4.2' -Description (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update -Version '5.4.2'))
            (New-FakeUpdate -Id $older -Title 'Contoso Tool 5.4.1' -Description (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update -Version '5.4.1')))
        $settings = @{} + $script:Settings
        $settings.ApprovalGroup = 'Pilot'
        $settings.DeclineSuperseded = $true
        $result = Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $settings
        $result.Outcome | Should -Be 'AlreadyPublished'
        $result.PackageId | Should -Be $current
        $result.Approved | Should -Be 'Pilot'
        Should -Invoke -ModuleName AppPackagerWsus New-WsusSoftwareDistributionPackageFile -Times 0 -Exactly
        $script:Approvals.ToArray() | Should -Be @('Pilot')
        $script:Declines.ToArray() | Should -Be @([string]$older)
    }

    It 'leaves an existing update that the server declined alone, with a warning' {
        $older = [guid]'00000000-0000-0000-0000-000000000541'
        $script:ServerUpdates = @(
            (New-FakeUpdate -Id ([guid]'00000000-0000-0000-0000-000000000542') -Title 'Contoso Tool 5.4.2' -Description (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update -Version '5.4.2') -Declined $true)
            (New-FakeUpdate -Id $older -Title 'Contoso Tool 5.4.1' -Description (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update -Version '5.4.1')))
        $settings = @{} + $script:Settings
        $settings.ApprovalGroup = 'Pilot'
        $settings.DeclineSuperseded = $true
        $result = Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $settings
        $result.Outcome | Should -Be 'AlreadyPublished'
        $script:Approvals.Count | Should -Be 0
        $script:Declines.Count | Should -Be 0
        @($result.Warnings)[0] | Should -Match 'declined on the server'
    }

    It 'publishes the same version under a new id when its earlier update is expired' {
        $expired = [guid]'00000000-0000-0000-0000-000000000542'
        $script:ServerUpdates = @(
            (New-FakeUpdate -Id $expired -Title 'Contoso Tool 5.4.2' -Description (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update -Version '5.4.2') -Declined $true) |
                Add-Member -NotePropertyName PublicationState -NotePropertyValue 'Expired' -PassThru)
        $result = Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $script:Settings
        $result.Outcome | Should -Be 'Published'
        $result.PackageId | Should -Not -Be $expired
        @($script:Plans[0].Plan.SupersededPackageIds).Count | Should -Be 0
    }

    It 'gives every new publish its own id' {
        [void](Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $script:Settings)
        [void](Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $script:Settings)
        $script:Plans.Count | Should -Be 2
        $script:Plans[0].Plan.PackageId | Should -Not -Be $script:Plans[1].Plan.PackageId
    }
    It 'keeps a successful publish when the approval or a decline fails afterwards' {
        $older = [guid]'00000000-0000-0000-0000-000000000541'
        $script:ServerUpdates = @(New-FakeUpdate -Id $older -Title 'Contoso Tool 5.4.1' -Description (Get-WsusIdentityLine -IdentityTag $script:Identity -PackageType Update -Version '5.4.1'))
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateApproval { throw "WSUS has no computer group named 'Pilto'." }
        $settings = @{} + $script:Settings
        $settings.ApprovalGroup = 'Pilto'
        $settings.DeclineSuperseded = $true
        $result = Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $settings
        $result.Outcome | Should -Be 'Published'
        $result.Approved | Should -Be ''
        @($result.Warnings)[0] | Should -Match "approval for Pilto failed: WSUS has no computer group named 'Pilto'"
        $script:Declines.ToArray() | Should -Be @([string]$older)
        $result.Message | Should -Match 'approval for Pilto failed'
    }

    It 'approves a new update for the configured group' {
        $settings = @{} + $script:Settings
        $settings.ApprovalGroup = 'Pilot'
        (Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $settings).Approved | Should -Be 'Pilot'
        $script:Approvals.ToArray() | Should -Be @('Pilot')
    }

    It 'refuses a blocking manifest before contacting the server' {
        { Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest -Overrides @{ InstallerType = 'MSIX' }) -ContentFolder $script:Content -Settings $script:Settings } |
            Should -Throw '*cannot be published to WSUS*InstallerTypeUnsupported*'
        Should -Invoke -ModuleName AppPackagerWsus Get-WsusServerConnection -Times 0 -Exactly
    }

    It 'marks a refusal so a caller can tell it apart from a failed server call' {
        $refusal = $null
        try { Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest -Overrides @{ InstallerType = 'MSIX'; InstallCommandLine = 'custom.cmd' }) -ContentFolder $script:Content -Settings $script:Settings }
        catch { $refusal = $_.Exception }
        $refusal | Should -Not -BeNullOrEmpty
        $refusal.Data.Contains('WsusRefusal') | Should -BeTrue
        @(([string]$refusal.Data['WsusRefusal']).Split(',')) | Should -Be @('InstallerTypeUnsupported', 'CustomInstall')

        Mock -ModuleName AppPackagerWsus Get-WsusServerConnection { throw 'The remote server returned an error: (503) Server Unavailable.' }
        $serverFault = $null
        try { Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $script:Settings }
        catch { $serverFault = $_.Exception }
        $serverFault | Should -Not -BeNullOrEmpty
        $serverFault.Data.Contains('WsusRefusal') | Should -BeFalse
    }

    It 'refuses a manifest that records a file outside the content folder before contacting the server' {
        $outside = Join-Path $script:Root 'outside-publish.txt'
        [System.IO.File]::WriteAllText($outside, 'marker')
        $manifest = New-HashedManifest
        $manifest.FileHashes = @($manifest.FileHashes) + [pscustomobject]@{ RelativePath = '..\outside-publish.txt'; Sha256 = (Get-FileHash -LiteralPath $outside -Algorithm SHA256).Hash; Size = 6 }
        { Publish-WsusSoftwareUpdate -Manifest $manifest -ContentFolder $script:Content -Settings $script:Settings } | Should -Throw '*PayloadOutsideStage*'
        Should -Invoke -ModuleName AppPackagerWsus Get-WsusServerConnection -Times 0 -Exactly
    }

    It 'refuses an installer whose bytes changed after staging' {
        $manifest = New-HashedManifest
        @($manifest.FileHashes | Where-Object RelativePath -eq 'setup.exe')[0].Sha256 = ('0' * 64)
        { Publish-WsusSoftwareUpdate -Manifest $manifest -ContentFolder $script:Content -Settings $script:Settings } | Should -Throw '*changed after it was staged*'
        Should -Invoke -ModuleName AppPackagerWsus Get-WsusServerConnection -Times 0 -Exactly
    }

    It 'refuses a server without a signing certificate and one whose certificate expired' {
        Mock -ModuleName AppPackagerWsus Get-WsusSigningCertificateObject { $null }
        { Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $script:Settings } | Should -Throw '*has no signing certificate*'
        Mock -ModuleName AppPackagerWsus Get-WsusSigningCertificateObject {
            [pscustomobject]@{ Subject = 'CN=x'; Issuer = 'CN=x'; Thumbprint = 'CD34'; NotBefore = (Get-Date).AddYears(-2); NotAfter = (Get-Date).AddDays(-1); PublicKey = $null }
        }
        { Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $script:Settings } | Should -Throw '*expired*'
        Should -Invoke -ModuleName AppPackagerWsus Invoke-WsusPublisher -Times 0 -Exactly
    }

    It 'passes MSI properties only and carries the transform with the other recorded files' {
        [System.IO.File]::WriteAllText((Join-Path $script:Content 'product.msi'), 'msi')
        $detection = [pscustomobject]@{ Type = 'RegistryKeyValue'; RegistryKeyRelative = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\7-Zip'; DisplayVersion = '24.08.00.0'; Is64Bit = $true }
        $manifest = New-HashedManifest -Overrides @{ InstallerType = 'MSI'; InstallerFile = 'product.msi'; InstallArgs = '/qn /norestart TRANSFORMS=custom.mst'; Detection = $detection }
        [void](Publish-WsusSoftwareUpdate -Manifest $manifest -ContentFolder $script:Content -Settings $script:Settings)
        $plan = $script:Plans[0].Plan
        $plan.CommandLine | Should -Be 'TRANSFORMS=custom.mst REBOOT=ReallySuppress'
        $plan.InstallerFile | Should -Be 'product.msi'
        @($script:Plans[0].Files | Sort-Object) | Should -Be @('custom.mst', 'product.msi', 'setup.exe')
    }

    It 'refuses an account that is not a WSUS administrator before building anything' {
        Mock -ModuleName AppPackagerWsus Get-WsusUserRoleName { 'Reporter' }
        { Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $script:Settings } | Should -Throw '*role Reporter*WSUS Administrators*'
        Should -Invoke -ModuleName AppPackagerWsus New-WsusSoftwareDistributionPackageFile -Times 0 -Exactly
    }

    It 'publishes when the certificate cannot be read, leaving the refusal to the server' {
        Mock -ModuleName AppPackagerWsus Get-WsusSigningCertificateObject { throw 'Access is denied' }
        (Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $script:Settings).Outcome | Should -Be 'Published'
    }

    It 'files every update under one vendor and one product category' {
        [void](Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest -Overrides @{ Publisher = 'Microsoft Corporation' }) -ContentFolder $script:Content -Settings $script:Settings)
        $script:Plans[0].Plan.VendorName | Should -Be 'AppPackager'
        $script:Plans[0].Plan.ProductName | Should -Be 'AppPackager Applications'
        $script:Plans[0].Plan.Description | Should -Match 'from Microsoft Corporation, published by AppPackager'
    }

    It 'keeps a long title under 80 characters with the version at the end' {
        $long = 'Contoso Enterprise Remote Management and Monitoring Agent for Windows Workstations (x64)'
        [void](Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest -Overrides @{ AppName = $long }) -ContentFolder $script:Content -Settings $script:Settings)
        $title = $script:Plans[0].Plan.Title
        $title.Length | Should -BeLessThan 80
        $title | Should -Match '\.\.\. 5\.4\.2$'
    }

    It 'names the version once in the description when the title already carries it' {
        [void](Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest -Overrides @{ AppName = 'Contoso Tool 5.4.2' }) -ContentFolder $script:Content -Settings $script:Settings)
        $script:Plans[0].Plan.Description | Should -Match '^Contoso Tool 5\.4\.2 from Contoso, published by AppPackager'
    }

    It 'adds the fix to a known server error' {
        Mock -ModuleName AppPackagerWsus Invoke-WsusPublisher { throw 'Exception occurred during publishing: Verification of file signature failed for file: \\wsus01\UpdateServicesPackages\x_1.cab' }
        { Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $script:Settings } |
            Should -Throw '*Verification of file signature failed*must trust the WSUS signing certificate*'
    }

    It 'names the server error when WSUS rejects the package and removes its working folder' {
        Mock -ModuleName AppPackagerWsus Invoke-WsusPublisher {
            $script:Published.Add([pscustomobject]@{ SdpPath = $SdpPath; SourceFolder = $SourceFolder })
            throw 'The signing certificate is not trusted'
        }
        { Publish-WsusSoftwareUpdate -Manifest (New-HashedManifest) -ContentFolder $script:Content -Settings $script:Settings } |
            Should -Throw '*WSUS rejected the package: The signing certificate is not trusted*'
        Test-Path -LiteralPath (Split-Path -Parent $script:Published[0].SourceFolder) | Should -BeFalse
    }
}

Describe 'Server content' {
    BeforeAll {
        $script:Settings = @{ ServerName = 'wsus01.contoso.com' }
        $script:PilotId = [guid]'11111111-2222-3333-4444-555555555555'
    }

    BeforeEach {
        Mock -ModuleName AppPackagerWsus Get-WsusServerConnection { [pscustomobject]@{ Name = 'wsus01.contoso.com' } }
        Mock -ModuleName AppPackagerWsus Get-WsusUserRoleName { 'Administrator' }
        Mock -ModuleName AppPackagerWsus Get-WsusComputerGroupObjects {
            @([pscustomobject]@{ Id = $script:PilotId; Name = 'Pilot' }, [pscustomobject]@{ Id = [guid]::NewGuid(); Name = 'All Computers' })
        }
        Mock -ModuleName AppPackagerWsus Remove-WsusPackageFolder { $true }
    }

    It 'lists AppPackager updates with their approval groups, newest first' {
        Mock -ModuleName AppPackagerWsus Get-WsusLocallyPublishedUpdateObjects {
            $mine = [pscustomobject]@{
                Id = [pscustomobject]@{ UpdateId = [guid]'aaaaaaaa-0000-0000-0000-000000000001' }; Title = 'Tool 2'; IsDeclined = $false; IsSuperseded = $false
                Description = (Get-WsusIdentityLine -IdentityTag 'AppPackager:catalog:package-tool/default' -PackageType Update); CreationDate = [datetime]'2026-09-20'
            }
            $mine | Add-Member -MemberType ScriptMethod -Name GetUpdateApprovals -Value { @([pscustomobject]@{ Action = 'Install'; ComputerTargetGroupId = $script:PilotId }) }
            $older = [pscustomobject]@{
                Id = [pscustomobject]@{ UpdateId = [guid]'aaaaaaaa-0000-0000-0000-000000000002' }; Title = 'Tool 1'; IsDeclined = $true; IsSuperseded = $true
                Description = (Get-WsusIdentityLine -IdentityTag 'AppPackager:catalog:package-tool/default' -PackageType Update); CreationDate = [datetime]'2026-08-20'
            }
            $older | Add-Member -MemberType ScriptMethod -Name GetUpdateApprovals -Value { @() }
            $foreign = [pscustomobject]@{
                Id = [pscustomobject]@{ UpdateId = [guid]::NewGuid() }; Title = 'Foreign'; IsDeclined = $false; IsSuperseded = $false
                Description = 'Published elsewhere'; CreationDate = [datetime]'2026-09-25'
            }
            $foreign | Add-Member -MemberType ScriptMethod -Name GetUpdateApprovals -Value { @() }
            @($older, $foreign, $mine)
        }
        $rows = @(Get-WsusPublishedUpdates -Settings $script:Settings)
        @($rows | ForEach-Object Title) | Should -Be @('Tool 2', 'Tool 1')
        $rows[0].ApprovedGroups | Should -Be 'Pilot'
        $rows[0].PackageType | Should -Be 'Update'
        $rows[1].Declined | Should -BeTrue
        @(Get-WsusPublishedUpdates -Settings $script:Settings -IncludeOtherPublishers).Count | Should -Be 3
    }

    It 'imports catalog updates by ID and reports failures without throwing' {
        Mock -ModuleName AppPackagerWsus Invoke-WsusCatalogImport { if ($UpdateId -eq [guid]'00000000-0000-0000-0000-00000000dead') { throw 'Update not found in the catalog' } }
        $ok = Import-WsusCatalogUpdate -Settings $script:Settings -UpdateId ([guid]'12345678-90ab-cdef-1234-567890abcdef')
        $ok.Ok | Should -BeTrue
        $bad = Import-WsusCatalogUpdate -Settings $script:Settings -UpdateId ([guid]'00000000-0000-0000-0000-00000000dead')
        $bad.Ok | Should -BeFalse
        $bad.Message | Should -Match 'not found in the catalog'
        $bad.Message | Should -Not -Match 'SchUseStrongCrypto'
        Should -Invoke -ModuleName AppPackagerWsus Invoke-WsusCatalogImport -Times 2 -Exactly
    }

    It 'names the WSUS server TLS 1.2 setting when the server cannot reach the catalog' {
        Mock -ModuleName AppPackagerWsus Invoke-WsusCatalogImport { throw 'The request was aborted: Could not create SSL/TLS secure channel.' }
        $result = Import-WsusCatalogUpdate -Settings $script:Settings -UpdateId ([guid]'8f7cd30d-371a-4646-9beb-16ca2cc57a2a')
        $result.Ok | Should -BeFalse
        $result.Message | Should -Match 'Could not create SSL/TLS secure channel'
        $result.Message | Should -Match 'On the WSUS server, set SchUseStrongCrypto = 1'
    }

    It 'reports an unreachable server as not connected instead of throwing' {
        Mock -ModuleName AppPackagerWsus Get-WsusServerConnection { throw 'WSUS server wsus01.contoso.com did not answer on port 8530 within 5 seconds.' }
        $status = Get-WsusServerStatus -Settings $script:Settings
        $status.Connected | Should -BeFalse
        $status.Message | Should -Match 'did not answer'
    }

    It 'declines before deleting, keeps the given order, and reports each update on its own' {
        $script:Calls = New-Object System.Collections.Generic.List[string]
        $old = [guid]'bbbbbbbb-0000-0000-0000-000000000001'
        $new = [guid]'bbbbbbbb-0000-0000-0000-000000000002'
        $missing = [guid]'bbbbbbbb-0000-0000-0000-000000000003'
        Mock -ModuleName AppPackagerWsus Get-WsusUpdateObject {
            if ($PackageId -eq [guid]'bbbbbbbb-0000-0000-0000-000000000001') { return [pscustomobject]@{ Title = 'Tool 1'; IsDeclined = $true } }
            if ($PackageId -eq [guid]'bbbbbbbb-0000-0000-0000-000000000002') { return [pscustomobject]@{ Title = 'Tool 2'; IsDeclined = $false } }
            return $null
        }
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateDecline { $script:Calls.Add('decline ' + $Update.Title) }
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateDeletion { $script:Calls.Add('delete ' + $UpdateId) }
        $results = @(Set-WsusPublishedUpdateState -Settings $script:Settings -PackageId @($old, $new, $missing) -Action Remove)
        $script:Calls.ToArray() | Should -Be @(('delete ' + $old), 'decline Tool 2', ('delete ' + $new))
        @($results | Where-Object Ok).Count | Should -Be 2
        ($results | Where-Object { $_.PackageId -eq $missing }).Message | Should -Match 'no update with id'
    }

    It 'explains a removal that WSUS refuses because another update references it' {
        Mock -ModuleName AppPackagerWsus Get-WsusUpdateObject { [pscustomobject]@{ Title = 'Tool 1'; IsDeclined = $true } }
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateDeletion { throw 'This update is still referenced by at least another update in the database.' }
        $result = @(Set-WsusPublishedUpdateState -Settings $script:Settings -PackageId ([guid]::NewGuid()) -Action Remove)[0]
        $result.Ok | Should -BeFalse
        $result.Message | Should -Match 'still referenced.*Expire this update instead, or remove the update that references it first'
        $result.Message | Should -Not -Match 'now declined'
    }

    It 'says so when a refused removal leaves the update declined' {
        Mock -ModuleName AppPackagerWsus Get-WsusUpdateObject { [pscustomobject]@{ Title = 'Tool 1'; IsDeclined = $false } }
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateDecline { }
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateDeletion { throw 'This update is still referenced by at least another update in the database.' }
        $result = @(Set-WsusPublishedUpdateState -Settings $script:Settings -PackageId ([guid]::NewGuid()) -Action Remove)[0]
        $result.Ok | Should -BeFalse
        $result.Message | Should -Match 'The update is now declined\.$'
    }

    It 'deletes the package folder after a removal and keeps the removal when the folder cannot go' {
        Mock -ModuleName AppPackagerWsus Get-WsusUpdateObject { [pscustomobject]@{ Title = 'Tool'; IsDeclined = $true; CreationDate = [datetime]'2026-09-01' } }
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateDeletion { }
        Mock -ModuleName AppPackagerWsus Remove-WsusPackageFolder { $true }
        $id = [guid]::NewGuid()
        (@(Set-WsusPublishedUpdateState -Settings $script:Settings -PackageId $id -Action Remove)[0]).Ok | Should -BeTrue
        Should -Invoke -ModuleName AppPackagerWsus Remove-WsusPackageFolder -Times 1 -Exactly -ParameterFilter { $PackageId -eq $id }
        Mock -ModuleName AppPackagerWsus Remove-WsusPackageFolder { throw 'Access is denied' }
        $result = @(Set-WsusPublishedUpdateState -Settings $script:Settings -PackageId $id -Action Remove)[0]
        $result.Ok | Should -BeTrue
        $result.Message | Should -Match 'package folder on the server was not deleted: Access is denied'
    }

    It 'explains a second expiry' {
        Mock -ModuleName AppPackagerWsus Get-WsusUpdateObject { [pscustomobject]@{ Title = 'Tool'; IsDeclined = $true; CreationDate = [datetime]'2026-09-01' } }
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateExpiry { throw 'Cannot revise or expire a package that has already been expired.' }
        (@(Set-WsusPublishedUpdateState -Settings $script:Settings -PackageId ([guid]::NewGuid()) -Action Expire)[0]).Message |
            Should -Match 'already been expired\. Only a locally published update that is not already expired'
    }

    It 'reads a missing signing certificate as none and rethrows any other failure' {
        $noCertificate = [pscustomobject]@{}
        $noCertificate | Add-Member -MemberType ScriptMethod -Name GetConfiguration -Value {
            $c = [pscustomobject]@{}
            $c | Add-Member -MemberType ScriptMethod -Name GetSigningCertificate -Value { param($p) throw (New-Object System.ComponentModel.Win32Exception(2)) }
            $c
        }
        InModuleScope AppPackagerWsus -Parameters @{ S = $noCertificate } { param($S) Get-WsusSigningCertificateObject -Server $S } | Should -BeNullOrEmpty
        $denied = [pscustomobject]@{}
        $denied | Add-Member -MemberType ScriptMethod -Name GetConfiguration -Value {
            $c = [pscustomobject]@{}
            $c | Add-Member -MemberType ScriptMethod -Name GetSigningCertificate -Value { param($p) throw (New-Object System.ComponentModel.Win32Exception(5)) }
            $c
        }
        { InModuleScope AppPackagerWsus -Parameters @{ S = $denied } { param($S) Get-WsusSigningCertificateObject -Server $S } } | Should -Throw
    }

    It 'expires, declines and approves through the same call' {
        Mock -ModuleName AppPackagerWsus Get-WsusUpdateObject { [pscustomobject]@{ Title = 'Tool'; IsDeclined = $false; CreationDate = [datetime]'2026-09-01' } }
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateExpiry { }
        Mock -ModuleName AppPackagerWsus Invoke-WsusUpdateApproval { }
        (@(Set-WsusPublishedUpdateState -Settings $script:Settings -PackageId ([guid]::NewGuid()) -Action Expire)[0]).Ok | Should -BeTrue
        (@(Set-WsusPublishedUpdateState -Settings $script:Settings -PackageId ([guid]::NewGuid()) -Action Approve -GroupName 'Pilot')[0]).Ok | Should -BeTrue
        Should -Invoke -ModuleName AppPackagerWsus Invoke-WsusUpdateExpiry -Times 1 -Exactly
        Should -Invoke -ModuleName AppPackagerWsus Invoke-WsusUpdateApproval -Times 1 -Exactly
        { Set-WsusPublishedUpdateState -Settings $script:Settings -PackageId ([guid]::NewGuid()) -Action Approve } | Should -Throw '*needs a computer group*'
    }

    It 'treats <Name> as this computer for the secure local connection' -TestCases @(
        @{ Name = 'localhost'; Expected = $true }
        @{ Name = '127.0.0.1'; Expected = $true }
        @{ Name = $env:COMPUTERNAME; Expected = $true }
        @{ Name = ($env:COMPUTERNAME.ToLowerInvariant() + '.'); Expected = $true }
        @{ Name = 'wsus01.contoso.com'; Expected = $false }
    ) {
        param($Name, $Expected)
        InModuleScope AppPackagerWsus -Parameters @{ N = $Name } { param($N) Test-WsusServerIsLocal -ServerName $N } | Should -Be $Expected
    }

    It 'refuses to send a PFX over a connection that is not secure' {
        $server = [pscustomobject]@{ Name = 'wsus01'; IsConnectionSecureForApiRemoting = $false }
        $server | Add-Member -MemberType ScriptMethod -Name GetConfiguration -Value { [pscustomobject]@{} }
        $password = New-Object System.Security.SecureString
        { InModuleScope AppPackagerWsus -Parameters @{ S = $server; P = $password } { param($S, $P) Set-WsusSigningCertificateCore -Server $S -PfxPath 'C:\x.pfx' -Password $P } } |
            Should -Throw '*not secure*Connect with SSL*'
    }

    It 'reports the account role with the server status' {
        Mock -ModuleName AppPackagerWsus Get-WsusServerConnection { [pscustomobject]@{ Name = 'wsus01.contoso.com'; Version = '10.0.26100.1'; IsConnectionSecureForApiRemoting = $true } }
        Mock -ModuleName AppPackagerWsus Get-WsusUserRoleName { 'Reporter' }
        Mock -ModuleName AppPackagerWsus Get-WsusSigningCertificateObject { $null }
        $status = Get-WsusServerStatus -Settings $script:Settings
        $status.Connected | Should -BeTrue
        $status.Role | Should -Be 'Reporter'
        $status.SecureConnection | Should -BeTrue
        $status.Message | Should -Match 'WSUS Administrators'
    }

    It 'refuses to approve an update for a group the server does not have' {
        $fake = [pscustomobject]@{ Title = 'Tool' }
        { InModuleScope AppPackagerWsus -Parameters @{ U = $fake } { param($U) Invoke-WsusUpdateApproval -Server ([pscustomobject]@{}) -Update $U -GroupName 'Nope' } } |
            Should -Throw "*no computer group named 'Nope'*"
    }
}

Describe 'Signing certificate import' {
    # Throwaway certificates live in Cert:\CurrentUser\My only for the export
    # and are removed with their keys; no trust store is touched.
    BeforeAll {
        $script:PfxPassword = ConvertTo-SecureString -String ('p-' + [guid]::NewGuid().ToString('N')) -AsPlainText -Force
        $script:PfxFiles = @{}
        foreach ($spec in @(
                @{ Name = 'Strong'; KeyLength = 2048; Type = 'CodeSigningCert' }
                @{ Name = 'Weak'; KeyLength = 1024; Type = 'CodeSigningCert' }
                @{ Name = 'Tls'; KeyLength = 2048; Type = 'SSLServerAuthentication' })) {
            $cert = New-SelfSignedCertificate -Type $spec.Type -Subject ('CN=AppPackager WSUS Test ' + $spec.Name) -KeyLength $spec.KeyLength `
                -KeyExportPolicy Exportable -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddDays(1)
            try {
                $path = Join-Path $script:Root ($spec.Name + '.pfx')
                [void](Export-PfxCertificate -Cert $cert -FilePath $path -Password $script:PfxPassword)
                $script:PfxFiles[$spec.Name] = $path
            }
            finally { Remove-Item -LiteralPath ('Cert:\CurrentUser\My\' + $cert.Thumbprint) -DeleteKey -Force -ErrorAction SilentlyContinue }
        }
    }

    BeforeEach {
        Mock -ModuleName AppPackagerWsus Get-WsusServerConnection { [pscustomobject]@{ Name = 'wsus01.contoso.com'; IsConnectionSecureForApiRemoting = $true } }
        Mock -ModuleName AppPackagerWsus Get-WsusUserRoleName { 'Administrator' }
        Mock -ModuleName AppPackagerWsus Set-WsusSigningCertificateCore { }
        Mock -ModuleName AppPackagerWsus Get-WsusServerStatus { [pscustomobject]@{ Connected = $true } }
    }

    It 'sets a 2048-bit code-signing certificate on the server' {
        [void](Set-WsusSigningCertificate -Settings @{ ServerName = 'wsus01.contoso.com' } -PfxPath $script:PfxFiles['Strong'] -Password $script:PfxPassword)
        Should -Invoke -ModuleName AppPackagerWsus Set-WsusSigningCertificateCore -Times 1 -Exactly
    }

    It 'refuses <Name> before the server configuration changes' -TestCases @(
        @{ Name = 'Weak'; Message = '*1024 bits*2048*' }
        @{ Name = 'Tls'; Message = '*not valid for code signing*' }
    ) {
        param($Name, $Message)
        { Set-WsusSigningCertificate -Settings @{ ServerName = 'wsus01.contoso.com' } -PfxPath $script:PfxFiles[$Name] -Password $script:PfxPassword } | Should -Throw $Message
        Should -Invoke -ModuleName AppPackagerWsus Set-WsusSigningCertificateCore -Times 0 -Exactly
    }

    It 'refuses a wrong password' {
        $wrong = ConvertTo-SecureString -String 'wrong' -AsPlainText -Force
        { Set-WsusSigningCertificate -Settings @{ ServerName = 'wsus01.contoso.com' } -PfxPath $script:PfxFiles['Strong'] -Password $wrong } | Should -Throw '*could not be opened*'
        Should -Invoke -ModuleName AppPackagerWsus Get-WsusServerConnection -Times 0 -Exactly
    }
}

Describe 'Install wrapper that maps an exit code' {
    It 'is reported as a step WSUS does not carry' {
        $mapped = Join-Path $TestDrive 'mapped.ps1'
        Set-Content -LiteralPath $mapped -Encoding ASCII -Value @(
            '$exePath = Join-Path $PSScriptRoot ''setup.exe''',
            '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
            'if ($proc.ExitCode -eq 111111) { exit 0 }',
            'exit $proc.ExitCode')
        $plain = Join-Path $TestDrive 'plain.ps1'
        Set-Content -LiteralPath $plain -Encoding ASCII -Value @(
            '$exePath = Join-Path $PSScriptRoot ''setup.exe''',
            'if (-not (Test-Path -LiteralPath $exePath)) { exit 0 }',
            '$proc = Start-Process -FilePath $exePath -ArgumentList @(''/S'') -Wait -PassThru -NoNewWindow',
            'exit $proc.ExitCode')
        InModuleScope AppPackagerWsus -Parameters @{ Mapped = $mapped; Plain = $plain } {
            @(Get-WsusInstallScriptExtras -Path $Mapped) -join ';' | Should -Match 'exit code mapped to success'
            @(Get-WsusInstallScriptExtras -Path $Plain).Count | Should -Be 0
        }
    }
}
