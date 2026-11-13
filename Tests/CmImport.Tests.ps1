#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
BeforeAll {
    $script:PackagersRoot = Join-Path $PSScriptRoot '..\Packagers'
    Import-Module (Join-Path $script:PackagersRoot 'AppPackagerCommon.psd1') -Force -DisableNameChecking
    Import-Module (Join-Path $script:PackagersRoot 'AppPackagerCmImport.psd1') -Force
    $script:FixtureRoot = Join-Path $PSScriptRoot 'Fixtures\ConfigMgrImport'

    function Read-Fixture {
        param([string]$Name)
        return (ConvertFrom-CMApplicationDigest -Xml (Get-Content -LiteralPath (Join-Path $script:FixtureRoot "$Name.app.xml") -Raw -Encoding UTF8))
    }

    function New-TestLeaf {
        param([string]$Source = 'File', [string]$Property = 'Exists', [string]$Operator = 'IsEquals', [string]$Expected = 'True', [string]$DataType = '',
              [string]$Path = '', [string]$Name = '', [string]$Hive = 'HKEY_LOCAL_MACHINE', [string]$Key = '', [bool]$Is64Bit = $true)
        [pscustomobject]@{ Kind = 'Leaf'; Source = $Source; Property = $Property; Operator = $Operator; Expected = $Expected; DataType = $DataType
            Path = $Path; Name = $Name; Hive = $Hive; Key = $Key; Is64Bit = $Is64Bit; Supported = $true; Reason = '' }
    }
    function New-TestGroup {
        param([string]$Operator, [object[]]$Children)
        [pscustomobject]@{ Kind = 'Group'; Operator = $Operator; Children = $Children }
    }

    function Invoke-TestScript {
        # Runs generated script text in Windows PowerShell, as the Intune client does.
        param([string]$Text)
        $file = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.ps1')
        [System.IO.File]::WriteAllText($file, $Text, (New-Object System.Text.UTF8Encoding($true)))
        $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $file 2>&1
        return [pscustomobject]@{ Exit = $LASTEXITCODE; Output = (($output | ForEach-Object { [string]$_ }) -join "`n").Trim() }
    }

    function Test-ParsesCleanly {
        param([string]$Text)
        $tokens = $null; $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tokens, [ref]$errors)
        return (@($errors).Count -eq 0)
    }

    $script:DecodeScript = { param($Rule) [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Rule.scriptContent)) }
}

Describe 'ConvertFrom-CMApplicationDigest' {
    It 'reads the application and an MSI deployment type' {
        $model = Read-Fixture 'single-msi'
        $model.Name | Should -Be 'IMPTEST Single MSI'
        $model.Publisher | Should -Be 'Contoso Lab'
        $model.Version | Should -Be '1.0'
        $model.LogicalName | Should -Match '^Application_'
        $model.DeploymentTypes.Count | Should -Be 1
        $dt = $model.DeploymentTypes[0]
        $dt.Technology | Should -Be 'MSI'
        $dt.Priority | Should -Be 1
        $dt.InstallCommand | Should -Be 'msiexec /i "32bitcompat.msi" /q'
        $dt.UninstallCommand | Should -Match '^msiexec /x \{4AF42476-4184-4857-AE0E-D8F4F794E34E\}'
        $dt.ContentLocation | Should -Be '\\cm01.contoso.com\ImpTest\msi'
        $dt.ContentFiles | Should -Contain '32bitcompat.msi'
        $dt.Detection.Method | Should -Be 'ProductCode'
        $dt.Detection.ProductCode | Should -Be '{4AF42476-4184-4857-AE0E-D8F4F794E34E}'
        $dt.ExecutionContext | Should -Be 'System'
    }

    It 'reads exit codes and run-time arguments' {
        $dt = (Read-Fixture 'multi-lang').DeploymentTypes[1]
        $dt.ExecutionContext | Should -Be 'System'
        $dt.MaxRuntimeMins | Should -Be 60
        $dt.EstimatedRuntimeMins | Should -Be 15
        @($dt.ExitCodes | Where-Object { $_.Class -eq 'SoftReboot' }).Code | Should -Be 3010
        @($dt.ExitCodes | Where-Object { $_.Class -eq 'FastRetry' }).Code | Should -Be 1618
    }

    It 'reads every clause type of an enhanced detection into an And tree' {
        $tree = (Read-Fixture 'clauses').DeploymentTypes[0].Detection.Tree
        $tree.Kind | Should -Be 'Group'
        $tree.Operator | Should -Be 'And'
        $tree.Children.Count | Should -Be 4
        $file, $registry, $key, $folder = $tree.Children
        $file.Source | Should -Be 'File'; $file.Property | Should -Be 'Version'; $file.Operator | Should -Be 'GreaterEquals'
        $file.Path | Should -Be 'C:\Program Files\ImpTest'; $file.Name | Should -Be 'app.exe'; $file.Is64Bit | Should -BeFalse; $file.Expected | Should -Be '1.2.3.4'
        $registry.Source | Should -Be 'Registry'; $registry.Key | Should -Be 'SOFTWARE\ImpTest'; $registry.Name | Should -Be 'Version'
        $registry.Is64Bit | Should -BeTrue; $registry.DataType | Should -Be 'Version'; $registry.Hive | Should -Be 'HKEY_LOCAL_MACHINE'
        $key.Source | Should -Be 'RegistryKey'; $key.Property | Should -Be 'Exists'; $key.Expected | Should -Be 'True'; $key.Is64Bit | Should -BeFalse
        $folder.Source | Should -Be 'Folder'; $folder.Property | Should -Be 'Exists'; $folder.Name | Should -Be 'data'
    }

    It 'reads Or and nested groups' {
        $tree = (Read-Fixture 'or-detect').DeploymentTypes[0].Detection.Tree
        $tree.Operator | Should -Be 'Or'
        $tree.Children.Count | Should -Be 2
        $tree.Children[0].Kind | Should -Be 'Leaf'
        $tree.Children[1].Kind | Should -Be 'Group'
        $tree.Children[1].Operator | Should -Be 'And'
        $tree.Children[1].Children.Count | Should -Be 2
    }

    It 'reads a script detection and a user-context deployment type' {
        $dt = (Read-Fixture 'script-detect').DeploymentTypes[0]
        $dt.Detection.Method | Should -Be 'Script'
        $dt.Detection.ScriptLanguage | Should -Be 'PowerShell'
        $dt.Detection.ScriptText | Should -Match 'notepad.exe'
        $dt.ExecutionContext | Should -Be 'User'
    }

    It 'keeps the priority order of several deployment types' {
        $dts = (Read-Fixture 'multi-lang').DeploymentTypes
        @($dts | ForEach-Object { $_.Name }) | Should -Be @('Lang - ja', 'Lang - x64', 'Lang - fallback')
        @($dts | ForEach-Object { $_.Priority }) | Should -Be @(1, 2, 3)
    }

    It 'reads operating system, language, memory and processor requirements' {
        $lang = (Read-Fixture 'multi-lang').DeploymentTypes
        $lang[0].Requirements[0].Kind | Should -Be 'Global'
        $lang[0].Requirements[0].LogicalName | Should -Be 'OSLanguage'
        $lang[0].Requirements[0].Operator | Should -Be 'OneOf'
        $lang[0].Requirements[0].Values | Should -Be @('1041')
        $lang[1].Requirements[0].Kind | Should -Be 'OsPlatform'
        $lang[1].Requirements[0].PlatformIds | Should -Be @('Windows/All_x64_Windows_10_higher')
        @($lang[2].Requirements).Count | Should -Be 0

        $misc = (Read-Fixture 'misc-reqs').DeploymentTypes[0].Requirements
        @($misc | ForEach-Object { $_.Kind + ':' + $_.LogicalName }) | Should -Be @('OsPlatform:', 'Global:TotalPhysicalMemory', 'Global:NumberOfProcessors')
        $misc[1].Values | Should -Be @('4294967296')
    }

    It 'reads dependencies and supersedence as references' {
        $dependent = (Read-Fixture 'dependent').DeploymentTypes[0]
        $dependent.Dependencies.Count | Should -Be 1
        $dependent.Dependencies[0].GroupName | Should -Be 'Needs arch'
        $dependent.Dependencies[0].Targets[0].DesiredState | Should -Be 'Required'
        $dependent.Dependencies[0].Targets[0].DeploymentTypeLogical | Should -Match '^DeploymentType_'
        $newer = (Read-Fixture 'newer').DeploymentTypes[0]
        $newer.Supersedes[0].Targets[0].DesiredState | Should -Be 'Prohibited'
    }

    It 'marks a technology that Win32 apps cannot carry' {
        $xml = (Get-Content -LiteralPath (Join-Path $script:FixtureRoot 'script-detect.app.xml') -Raw -Encoding UTF8) -replace '<Technology>Script</Technology>', '<Technology>AppV5X</Technology>'
        $dt = (ConvertFrom-CMApplicationDigest -Xml $xml).DeploymentTypes[0]
        $dt.UnsupportedReason | Should -Match 'AppV5X'
    }

    It 'refuses a document that is not an application digest' {
        { ConvertFrom-CMApplicationDigest -Xml '<root/>' } | Should -Throw '*AppMgmtDigest*'
    }
}

Describe 'ConvertTo-IntuneDetectionRules' {
    It 'maps an MSI product code to a product code rule' {
        $result = ConvertTo-IntuneDetectionRules -Detection (Read-Fixture 'single-msi').DeploymentTypes[0].Detection
        $result.Mode | Should -Be 'ProductCode'
        $result.Rules[0]['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppProductCodeRule'
        $result.Rules[0]['productCode'] | Should -Be '{4AF42476-4184-4857-AE0E-D8F4F794E34E}'
        $result.Rules[0]['ruleType'] | Should -Be 'detection'
    }

    It 'maps a flat And of file, registry and folder clauses to native rules with the right view flags' {
        $result = ConvertTo-IntuneDetectionRules -Detection (Read-Fixture 'clauses').DeploymentTypes[0].Detection
        $result.Mode | Should -Be 'Native'
        $result.Rules.Count | Should -Be 4
        $file, $registry, $key, $folder = $result.Rules
        $file['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppFileSystemRule'
        $file['operationType'] | Should -Be 'version'; $file['operator'] | Should -Be 'greaterThanOrEqual'; $file['comparisonValue'] | Should -Be '1.2.3.4'
        $file['check32BitOn64System'] | Should -BeTrue
        $registry['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppRegistryRule'
        $registry['keyPath'] | Should -Be 'HKEY_LOCAL_MACHINE\SOFTWARE\ImpTest'; $registry['valueName'] | Should -Be 'Version'
        $registry['operationType'] | Should -Be 'version'; $registry['check32BitOn64System'] | Should -BeFalse
        $key['operationType'] | Should -Be 'exists'; $key['check32BitOn64System'] | Should -BeTrue; $key['valueName'] | Should -BeNullOrEmpty
        $folder['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppFileSystemRule'
        $folder['operationType'] | Should -Be 'exists'; $folder['fileOrFolderName'] | Should -Be 'data'
    }

    It 'carries a PowerShell detection script as one script rule' {
        $result = ConvertTo-IntuneDetectionRules -Detection (Read-Fixture 'script-detect').DeploymentTypes[0].Detection
        $result.Mode | Should -Be 'Script'
        $result.Rules.Count | Should -Be 1
        $result.Rules[0]['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppPowerShellScriptRule'
        (& $script:DecodeScript $result.Rules[0]) | Should -Match 'notepad.exe'
        $result.Rules[0]['enforceSignatureCheck'] | Should -BeFalse
    }

    It 'turns Or and groups into one script rule, with a review finding' {
        $result = ConvertTo-IntuneDetectionRules -Detection (Read-Fixture 'or-detect').DeploymentTypes[0].Detection
        $result.Mode | Should -Be 'ExpressionScript'
        $result.Rules.Count | Should -Be 1
        $result.Findings.Code | Should -Contain 'DetectionConvertedToScript'
        Test-ParsesCleanly (& $script:DecodeScript $result.Rules[0]) | Should -BeTrue
    }

    It 'refuses the script conversion when native rules only are requested' {
        $result = ConvertTo-IntuneDetectionRules -Detection (Read-Fixture 'or-detect').DeploymentTypes[0].Detection -NativeOnly
        $result.Rules.Count | Should -Be 0
        ($result.Findings | Where-Object { $_.Severity -eq 'Blocking' }).Code | Should -Be 'DetectionNeedsScript'
    }

    It 'moves a prefix comparison and a wildcard file name to the script' {
        $prefix = New-TestLeaf -Source Registry -Property Value -Operator BeginsWith -Expected 'Win' -DataType String -Key 'SOFTWARE\X' -Name 'N'
        $result = ConvertTo-IntuneDetectionRules -Detection ([pscustomobject]@{ Method = 'Enhanced'; Tree = $prefix })
        $result.Mode | Should -Be 'ExpressionScript'
        $wild = New-TestLeaf -Source File -Property Exists -Path 'C:\X' -Name '*.dll'
        (ConvertTo-IntuneDetectionRules -Detection ([pscustomobject]@{ Method = 'Enhanced'; Tree = $wild })).Mode | Should -Be 'ExpressionScript'
    }

    It 'blocks a clause it cannot read' {
        $bad = [pscustomobject]@{ Kind = 'Leaf'; Source = 'WQL'; Supported = $false; Reason = 'not converted' }
        $result = ConvertTo-IntuneDetectionRules -Detection ([pscustomobject]@{ Method = 'Enhanced'; Tree = $bad })
        ($result.Findings | Where-Object { $_.Severity -eq 'Blocking' }).Code | Should -Be 'DetectionNotConvertible'
    }
}

Describe 'New-IntuneExpressionDetectionScript' {
    It 'exits 0 with output when an Or has one true branch' {
        $true1 = New-TestLeaf -Source File -Property Exists -Expected 'True' -Path '%SystemRoot%\System32' -Name 'kernel32.dll'
        $false1 = New-TestLeaf -Source File -Property Exists -Expected 'True' -Path '%SystemRoot%\System32' -Name 'no-such-file-imptest.dll'
        $run = Invoke-TestScript (New-IntuneExpressionDetectionScript -Tree (New-TestGroup 'Or' @($false1, $true1)))
        $run.Exit | Should -Be 0
        $run.Output | Should -Be 'Detected'
    }

    It 'exits 1 when an And has one false branch' {
        $true1 = New-TestLeaf -Source File -Property Exists -Expected 'True' -Path '%SystemRoot%\System32' -Name 'kernel32.dll'
        $false1 = New-TestLeaf -Source File -Property Exists -Expected 'True' -Path '%SystemRoot%\System32' -Name 'no-such-file-imptest.dll'
        $run = Invoke-TestScript (New-IntuneExpressionDetectionScript -Tree (New-TestGroup 'And' @($true1, $false1)))
        $run.Exit | Should -Be 1
        $run.Output | Should -BeNullOrEmpty
    }

    It 'evaluates nested groups, absence, registry values and versions' {
        $present = New-TestLeaf -Source RegistryKey -Property Exists -Expected 'True' -Key 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $absent = New-TestLeaf -Source RegistryKey -Property Exists -Expected 'False' -Key 'SOFTWARE\ImpTestNoSuchKey'
        $build = New-TestLeaf -Source Registry -Property Value -Operator GreaterEquals -Expected '1' -DataType Int64 -Key 'SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'CurrentBuildNumber'
        $name = New-TestLeaf -Source Registry -Property Value -Operator BeginsWith -Expected 'Windows' -DataType String -Key 'SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'ProductName'
        $version = New-TestLeaf -Source File -Property Version -Operator GreaterEquals -Expected '6.0.0.0' -DataType Version -Path '%SystemRoot%\System32' -Name 'kernel32.dll'
        $tree = New-TestGroup 'And' @($present, $absent, (New-TestGroup 'Or' @($build, $name)), $version)
        $run = Invoke-TestScript (New-IntuneExpressionDetectionScript -Tree $tree)
        $run.Exit | Should -Be 0
        $tooNew = New-TestLeaf -Source Registry -Property Value -Operator GreaterEquals -Expected '99999999' -DataType Int64 -Key 'SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'CurrentBuildNumber'
        (Invoke-TestScript (New-IntuneExpressionDetectionScript -Tree $tooNew)).Exit | Should -Be 1
    }

    It 'reads the 32-bit registry view when the clause is 32-bit' {
        # HKLM\SOFTWARE\WOW6432Node\Microsoft exists on 64-bit Windows; the 32-bit view of SOFTWARE\Microsoft is that key.
        $wow = New-TestLeaf -Source RegistryKey -Property Exists -Expected 'True' -Key 'SOFTWARE\Microsoft' -Is64Bit $false
        (Invoke-TestScript (New-IntuneExpressionDetectionScript -Tree $wow)).Exit | Should -Be 0
        $missing32 = New-TestLeaf -Source RegistryKey -Property Exists -Expected 'True' -Key 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\ImpTestNoSuchSubKey' -Is64Bit $false
        (Invoke-TestScript (New-IntuneExpressionDetectionScript -Tree $missing32)).Exit | Should -Be 1
    }

    It 'quotes values with single quotes' {
        $leaf = New-TestLeaf -Source File -Property Exists -Path "C:\O'Brien" -Name "a'b.txt"
        $text = New-IntuneExpressionDetectionScript -Tree $leaf
        Test-ParsesCleanly $text | Should -BeTrue
        $text | Should -Match "O''Brien"
    }
}

Describe 'ConvertTo-IntuneRequirementModel' {
    It 'takes architecture, memory and processors natively' {
        $model = ConvertTo-IntuneRequirementModel -Requirements @((Read-Fixture 'misc-reqs').DeploymentTypes[0].Requirements)
        $model.Architectures | Should -Be @('x64')
        $model.Native['MinimumMemoryInMB'] | Should -Be 4096
        $model.Native['MinimumNumberOfProcessors'] | Should -Be 2
        $model.ScriptOnly.Count | Should -Be 0
        $model.Evaluable | Should -BeTrue
        $model.HasRules | Should -BeTrue
    }

    It 'sends a language rule to the script' {
        $model = ConvertTo-IntuneRequirementModel -Requirements @((Read-Fixture 'multi-lang').DeploymentTypes[0].Requirements)
        $model.Predicates.Count | Should -Be 1
        $model.ScriptOnly.Count | Should -Be 1
        $model.ScriptOnly[0].Type | Should -Be 'Language'
        $model.ScriptOnly[0].Values | Should -Be @(1041)
        $model.Architectures | Should -BeNullOrEmpty
    }

    It 'reports a deployment type without rules' {
        $model = ConvertTo-IntuneRequirementModel -Requirements @()
        $model.HasRules | Should -BeFalse
        $model.Evaluable | Should -BeTrue
    }

    It 'sends a Windows 11 floor and a NoneOf rule to the script' {
        $win11 = [pscustomobject]@{ Kind = 'OsPlatform'; DisplayName = 'w11'; Operator = 'OneOf'; PlatformIds = @('Windows/All_x64_Windows_11_higher') }
        $none = [pscustomobject]@{ Kind = 'OsPlatform'; DisplayName = 'not arm'; Operator = 'NoneOf'; PlatformIds = @('Windows/All_ARM64_Windows_10_higher') }
        (ConvertTo-IntuneRequirementModel -Requirements @($win11)).ScriptOnly.Count | Should -Be 1
        (ConvertTo-IntuneRequirementModel -Requirements @($none)).ScriptOnly.Count | Should -Be 1
    }

    It 'lists a rule it cannot evaluate as unrecognized' {
        $custom = [pscustomobject]@{ Kind = 'Unsupported'; DisplayName = 'Custom thing'; Operator = ''; PlatformIds = @(); LogicalName = ''; Values = @(); DataType = ''; Reason = 'custom' }
        $disk = [pscustomobject]@{ Kind = 'Global'; DisplayName = 'Disk'; Operator = 'GreaterEquals'; PlatformIds = @(); LogicalName = 'FreeDiskSpace'; Values = @('5'); DataType = 'Int64'; Reason = '' }
        $model = ConvertTo-IntuneRequirementModel -Requirements @($custom, $disk)
        $model.Evaluable | Should -BeFalse
        $model.Unrecognized.Count | Should -Be 2
    }

    It 'refuses an operating system the Win32 app type cannot reach' {
        $old = [pscustomobject]@{ Kind = 'OsPlatform'; DisplayName = 'Win7'; Operator = 'OneOf'; PlatformIds = @('Windows/All_x64_Windows_7_Client') }
        (ConvertTo-IntuneRequirementModel -Requirements @($old)).Evaluable | Should -BeFalse
    }
}

Describe 'Get-CMImportVariantExclusion' {
    BeforeAll {
        function Get-RequirementModel([string]$Fixture, [int]$Index) { ConvertTo-IntuneRequirementModel -Requirements @((Read-Fixture $Fixture).DeploymentTypes[$Index].Requirements) }
    }

    It 'needs nothing for the first variant' {
        (Get-CMImportVariantExclusion -Own (Get-RequirementModel 'multi-arch' 0) -Earlier @()).Status | Should -Be 'None'
    }

    It 'needs no script when the architectures are disjoint' {
        $result = Get-CMImportVariantExclusion -Own (Get-RequirementModel 'multi-arch' 1) -Earlier @((Get-RequirementModel 'multi-arch' 0))
        $result.Status | Should -Be 'Native'
    }

    It 'needs a script when an earlier variant is not disjoint' {
        $result = Get-CMImportVariantExclusion -Own (Get-RequirementModel 'multi-lang' 1) -Earlier @((Get-RequirementModel 'multi-lang' 0))
        $result.Status | Should -Be 'Script'
        $result.EarlierSets.Count | Should -Be 1
        $fallback = Get-CMImportVariantExclusion -Own (Get-RequirementModel 'multi-lang' 2) -Earlier @((Get-RequirementModel 'multi-lang' 0), (Get-RequirementModel 'multi-lang' 1))
        $fallback.Status | Should -Be 'Script'
        $fallback.EarlierSets.Count | Should -Be 2
    }

    It 'reports a variant that an unconditional earlier variant shadows' {
        $result = Get-CMImportVariantExclusion -Own (Get-RequirementModel 'unreachable' 1) -Earlier @((Get-RequirementModel 'unreachable' 0))
        $result.Status | Should -Be 'Unreachable'
    }

    It 'reports an exclusion it cannot prove' {
        $custom = [pscustomobject]@{ Kind = 'Unsupported'; DisplayName = 'Custom'; Operator = ''; PlatformIds = @(); LogicalName = ''; Values = @(); DataType = ''; Reason = 'custom' }
        $earlier = ConvertTo-IntuneRequirementModel -Requirements @($custom)
        (Get-CMImportVariantExclusion -Own (Get-RequirementModel 'multi-arch' 0) -Earlier @($earlier)).Status | Should -Be 'Unprovable'
    }
}

Describe 'New-IntuneVariantRequirementScript' {
    BeforeAll {
        $script:CurrentLanguage = [int](Get-CimInstance -ClassName Win32_OperatingSystem).OSLanguage
        $script:CurrentIsArm = ([int](Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1).Architecture -eq 12)
        function New-ArchPredicate([string]$Arch) { [pscustomobject]@{ Type = 'OsPlatform'; Negate = $false; Pairs = @([pscustomobject]@{ Architecture = $Arch; MinimumBuild = 10240 }) } }
        function New-LanguagePredicate([int[]]$Lcids) { [pscustomobject]@{ Type = 'Language'; Comparison = '-in'; Values = $Lcids } }
    }

    It 'prints false when an earlier variant applies to this device' {
        $arch = $(if ($script:CurrentIsArm) { 'arm64' } else { 'x64' })
        $run = Invoke-TestScript (New-IntuneVariantRequirementScript -EarlierSets @(, @((New-ArchPredicate $arch))))
        $run.Exit | Should -Be 0
        $run.Output | Should -Be 'false'
    }

    It 'prints true when no earlier variant applies' {
        $run = Invoke-TestScript (New-IntuneVariantRequirementScript -EarlierSets @(, @((New-ArchPredicate 'x86'))))
        $run.Output | Should -Be 'true'
        $run = Invoke-TestScript (New-IntuneVariantRequirementScript -EarlierSets @(, @((New-LanguagePredicate @(1)))))
        $run.Output | Should -Be 'true'
    }

    It 'prints false for the language of this device and checks every earlier set' {
        $run = Invoke-TestScript (New-IntuneVariantRequirementScript -EarlierSets @(, @((New-LanguagePredicate @($script:CurrentLanguage)))))
        $run.Output | Should -Be 'false'
        (Invoke-TestScript (New-IntuneVariantRequirementScript -EarlierSets @(@((New-ArchPredicate 'x86')), @((New-LanguagePredicate @($script:CurrentLanguage)))))).Output | Should -Be 'false'
    }

    It 'combines the own predicates with the exclusion' {
        $own = @((New-LanguagePredicate @($script:CurrentLanguage)))
        (Invoke-TestScript (New-IntuneVariantRequirementScript -OwnPredicates $own -EarlierSets @(, @((New-ArchPredicate 'x86'))))).Output | Should -Be 'true'
        $otherLanguage = @((New-LanguagePredicate @(1)))
        (Invoke-TestScript (New-IntuneVariantRequirementScript -OwnPredicates $otherLanguage)).Output | Should -Be 'false'
    }

    It 'prints true when nothing is required' {
        (Invoke-TestScript (New-IntuneVariantRequirementScript)).Output | Should -Be 'true'
    }
}

Describe 'Get-ImpSetupFile and ConvertTo-IntuneReturnCodes' {
    It 'finds the file that the command runs' {
        Get-ImpSetupFile -InstallCommand 'install.bat' -ContentFiles @('install.bat', 'uninstall.bat') | Should -Be 'install.bat'
        Get-ImpSetupFile -InstallCommand 'msiexec /i "Tool Setup.msi" /qn' -ContentFiles @('Tool Setup.msi', 'a.txt') | Should -Be 'Tool Setup.msi'
        Get-ImpSetupFile -InstallCommand '"setup.exe" /S' -ContentFiles @('SETUP.EXE') | Should -Be 'SETUP.EXE'
        Get-ImpSetupFile -InstallCommand 'powershell.exe -File .\run.ps1' -ContentFiles @('run.ps1') | Should -Be 'run.ps1'
        Get-ImpSetupFile -InstallCommand 'cmd /c something' -ContentFiles @('readme.txt', 'x.exe') | Should -Be 'x.exe'
        Get-ImpSetupFile -InstallCommand 'cmd /c something' -ContentFiles @('readme.txt') | Should -Be ''
    }

    It 'maps the exit code classes and supplies defaults' {
        $codes = ConvertTo-IntuneReturnCodes -ExitCodes @([pscustomobject]@{ Code = 0; Class = 'Success' }, [pscustomobject]@{ Code = 3010; Class = 'SoftReboot' }, [pscustomobject]@{ Code = 1641; Class = 'HardReboot' }, [pscustomobject]@{ Code = 1618; Class = 'FastRetry' }, [pscustomobject]@{ Code = 5; Class = 'Failure' })
        @($codes | ForEach-Object { "{0}={1}" -f $_.returnCode, $_.type }) | Should -Be @('0=success', '3010=softReboot', '1641=hardReboot', '1618=retry', '5=failed')
        (ConvertTo-IntuneReturnCodes -ExitCodes @()).Count | Should -Be 5
    }
}

Describe 'New-CMImportPlan' {
    It 'plans a single deployment type under the application name' {
        $plan = New-CMImportPlan -Model (Read-Fixture 'clauses')
        $plan.DeploymentTypes.Count | Should -Be 1
        $dt = $plan.DeploymentTypes[0]
        $dt.Convert | Should -BeTrue
        $dt.IntuneName | Should -Be 'IMPTEST Clauses'
        $dt.Manifest.AppName | Should -Be 'IMPTEST Clauses'
        $dt.Manifest.SetupFile | Should -Be 'install.bat'
        $dt.Manifest.InstallCommandLine | Should -Be 'install.bat'
        $dt.Manifest.InstallationBehaviorType | Should -Be 'InstallForSystem'
        $dt.Manifest.IntuneRules.Count | Should -Be 4
        $dt.Manifest.IntuneAllowedArchitectures | Should -Be 'x86,x64'
        $dt.Manifest.ApplicationId | Should -Match '^cm:Application_'
        $dt.Manifest.ProfileId | Should -Match '^DeploymentType_'
        $dt.Manifest.Timing.MaximumRuntimeMins | Should -Be 45
        $dt.Manifest.Timing.IntuneApplied | Should -BeFalse
        $plan.Findings.Count | Should -Be 0
    }

    It 'maps a user-context deployment type and keeps the interaction flag only for system installs' {
        $dt = (New-CMImportPlan -Model (Read-Fixture 'script-detect')).DeploymentTypes[0]
        $dt.Manifest.InstallationBehaviorType | Should -Be 'InstallForUser'
        $dt.Manifest.RequireUserInteraction | Should -BeFalse
        $dt.Findings.Code | Should -Contain 'UserInteractionPossible'
        $xml = (Get-Content -LiteralPath (Join-Path $script:FixtureRoot 'clauses.app.xml') -Raw -Encoding UTF8) -replace '(<Arg Name="RequiresUserInteraction" Type="Boolean">)false', '${1}true'
        $blocked = (New-CMImportPlan -Model (ConvertFrom-CMApplicationDigest -Xml $xml)).DeploymentTypes[0]
        $blocked.Manifest.RequireUserInteraction | Should -BeTrue
        @(Get-IntuneCompatibilityFindings -Manifest $blocked.Manifest | Where-Object { $_.Code -eq 'InteractiveInstall' }).Count | Should -Be 1
    }

    It 'makes one Intune app per deployment type, named with the deployment type' {
        $plan = New-CMImportPlan -Model (Read-Fixture 'multi-arch')
        @($plan.DeploymentTypes.IntuneName) | Should -Be @('IMPTEST Multi Arch - Multi - x64', 'IMPTEST Multi Arch - Multi - x86')
        $x64, $x86 = $plan.DeploymentTypes
        $x64.Manifest.IntuneAllowedArchitectures | Should -Be 'x64'
        $x86.Manifest.IntuneAllowedArchitectures | Should -Be 'x86'
        $x64.Manifest.ProfileId | Should -Not -Be $x86.Manifest.ProfileId
        $x64.Manifest.ApplicationId | Should -Be $x86.Manifest.ApplicationId
        $x64.Manifest.SetupFile | Should -Be 'install.bat'
        @($x86.Manifest.IntuneRules | Where-Object { $_['ruleType'] -eq 'requirement' }).Count | Should -Be 0
        $plan.Findings.Code | Should -Contain 'VariantApps'
    }

    It 'adds a requirement script to keep first match wins' {
        $plan = New-CMImportPlan -Model (Read-Fixture 'multi-lang')
        $ja, $x64, $fallback = $plan.DeploymentTypes
        @($ja.Manifest.IntuneRules | Where-Object { $_['ruleType'] -eq 'requirement' }).Count | Should -Be 1
        $requirement = @($x64.Manifest.IntuneRules | Where-Object { $_['ruleType'] -eq 'requirement' })[0]
        $requirement['operationType'] | Should -Be 'boolean'
        $requirement['operator'] | Should -Be 'equal'
        $requirement['comparisonValue'] | Should -Be 'true'
        (& $script:DecodeScript $requirement) | Should -Match '1041'
        $fallbackRule = @($fallback.Manifest.IntuneRules | Where-Object { $_['ruleType'] -eq 'requirement' })[0]
        (& $script:DecodeScript $fallbackRule) | Should -Match 'x64'
    }

    It 'leaves out a deployment type that ConfigMgr can never choose' {
        $plan = New-CMImportPlan -Model (Read-Fixture 'unreachable')
        $plan.DeploymentTypes[0].Convert | Should -BeTrue
        $plan.DeploymentTypes[1].Convert | Should -BeFalse
        $plan.DeploymentTypes[1].Findings.Code | Should -Contain 'UnreachableDeploymentType'
        $plan.DeploymentTypes[1].Manifest | Should -BeNullOrEmpty
    }

    It 'blocks an unprovable exclusion unless the strategy is AssignByGroup' {
        $xml = (Get-Content -LiteralPath (Join-Path $script:FixtureRoot 'multi-arch.app.xml') -Raw -Encoding UTF8)
        $xml = $xml -replace '<OperatingSystemExpression>\s*<Operator>OneOf</Operator>\s*<Operands>\s*<RuleExpression RuleId="Windows/All_x64_Windows_10_higher"\s*/>\s*</Operands>\s*</OperatingSystemExpression>', '<Foo/>'
        $model = ConvertFrom-CMApplicationDigest -Xml $xml
        $model.DeploymentTypes[0].Requirements[0].Kind | Should -Be 'Unsupported'
        $blocked = New-CMImportPlan -Model $model
        $blocked.DeploymentTypes[1].Convert | Should -BeFalse
        $blocked.DeploymentTypes[1].Findings.Code | Should -Contain 'ExclusivityUnprovable'
        $byGroup = New-CMImportPlan -Model $model -VariantStrategy AssignByGroup
        $byGroup.DeploymentTypes[1].Convert | Should -BeTrue
        $byGroup.DeploymentTypes[1].Findings.Code | Should -Contain 'AssignByGroupRequired'
    }

    It 'does not add exclusion scripts under AssignByGroup' {
        $plan = New-CMImportPlan -Model (Read-Fixture 'multi-lang') -VariantStrategy AssignByGroup
        @($plan.DeploymentTypes[2].Manifest.IntuneRules | Where-Object { $_['ruleType'] -eq 'requirement' }).Count | Should -Be 0
        $plan.DeploymentTypes[1].Findings.Code | Should -Contain 'AssignByGroupRequired'
    }

    It 'lists dependencies and supersedence without creating them' {
        $dependent = New-CMImportPlan -Model (Read-Fixture 'dependent')
        $dependent.Relationships.Count | Should -Be 1
        $dependent.Relationships[0].Type | Should -Be 'Dependency'
        $dependent.Relationships[0].DesiredState | Should -Be 'Required'
        $dependent.Findings.Code | Should -Contain 'RelationshipsNotCreated'
        $newer = New-CMImportPlan -Model (Read-Fixture 'newer')
        $newer.Relationships[0].Type | Should -Be 'Supersedence'
    }

    It 'takes minimum hardware values natively and keeps the return codes' {
        $dt = (New-CMImportPlan -Model (Read-Fixture 'misc-reqs')).DeploymentTypes[0]
        $dt.Manifest.IntuneMinimums['MinimumMemoryInMB'] | Should -Be 4096
        $dt.Manifest.IntuneMinimums['MinimumNumberOfProcessors'] | Should -Be 2
        @($dt.Manifest.IntuneReturnCodes | Where-Object { $_.type -eq 'softReboot' }).returnCode | Should -Be 3010
        $dt.Manifest.IntuneRestartBehavior | Should -Be 'basedOnReturnCode'
    }

    It 'does not convert a technology Win32 apps cannot carry' {
        $xml = (Get-Content -LiteralPath (Join-Path $script:FixtureRoot 'script-detect.app.xml') -Raw -Encoding UTF8) -replace '<Technology>Script</Technology>', '<Technology>AppV5X</Technology>'
        $dt = (New-CMImportPlan -Model (ConvertFrom-CMApplicationDigest -Xml $xml)).DeploymentTypes[0]
        $dt.Convert | Should -BeFalse
        $dt.Findings.Code | Should -Contain 'TechnologyNotConverted'
    }

    It 'does not convert a deployment type whose content holds no setup file' {
        $xml = (Get-Content -LiteralPath (Join-Path $script:FixtureRoot 'clauses.app.xml') -Raw -Encoding UTF8) -replace '<File Name="install.bat" Size="\d+"\s*/>', '' -replace '<File Name="uninstall.bat" Size="\d+"\s*/>', ''
        $dt = (New-CMImportPlan -Model (ConvertFrom-CMApplicationDigest -Xml $xml)).DeploymentTypes[0]
        $dt.Convert | Should -BeFalse
        $dt.Findings.Code | Should -Contain 'SetupFileNotFound'
    }

    It 'uses a command that does nothing when ConfigMgr stores no uninstall' {
        $xml = (Get-Content -LiteralPath (Join-Path $script:FixtureRoot 'clauses.app.xml') -Raw -Encoding UTF8) -replace '<UninstallSetting>SameAsInstall</UninstallSetting>', '<UninstallSetting>NoneRequired</UninstallSetting>'
        $dt = (New-CMImportPlan -Model (ConvertFrom-CMApplicationDigest -Xml $xml)).DeploymentTypes[0]
        $dt.Manifest.UninstallCommandLine | Should -Be 'cmd.exe /c exit 0'
        $dt.Findings.Code | Should -Contain 'NoUninstallCommand'
    }

    It 'writes a readable report' {
        $text = Format-CMImportReport -Plan (New-CMImportPlan -Model (Read-Fixture 'multi-lang'))
        $text | Should -Match 'Deployment type 1: Lang - ja'
        $text | Should -Match 'Deployment type 3: Lang - fallback'
        $text | Should -Match 'VariantExclusionScript'
    }
}

Describe 'Export-CMImportStage' {
    BeforeEach {
        $script:Source = Join-Path $TestDrive 'source'
        $script:Out = Join-Path $TestDrive 'out'
        foreach ($d in 'x64', 'x86', 'fallback', 'ja') {
            New-Item -ItemType Directory -Force -Path (Join-Path $script:Source $d) | Out-Null
            Set-Content -LiteralPath (Join-Path $script:Source "$d\install.bat") -Value 'exit /b 0'
            Set-Content -LiteralPath (Join-Path $script:Source "$d\uninstall.bat") -Value 'exit /b 0'
        }
    }

    It 'copies content per deployment type and writes the manifests and reports' {
        $plan = New-CMImportPlan -Model (Read-Fixture 'multi-arch')
        foreach ($dt in $plan.DeploymentTypes) { $dt.ContentLocation = Join-Path $script:Source $(if ($dt.Priority -eq 1) { 'x64' } else { 'x86' }) }
        $stage = Export-CMImportStage -Plan $plan -OutputRoot $script:Out
        $stage.Staged.Count | Should -Be 2
        foreach ($item in $stage.Staged) {
            Test-Path -LiteralPath (Join-Path $item.ContentFolder 'install.bat') | Should -BeTrue
            $manifest = Get-Content -LiteralPath $item.ManifestPath -Raw | ConvertFrom-Json
            $manifest.AppName | Should -Be $item.IntuneName
            $manifest.ContentRoot | Should -Be $item.ContentFolder
            @($manifest.IntuneRules).Count | Should -BeGreaterThan 0
        }
        foreach ($name in 'conversion-report.txt', 'conversion-plan.json', 'relationships.json') {
            Test-Path -LiteralPath (Join-Path $stage.ApplicationFolder $name) | Should -BeTrue
        }
    }

    It 'mirrors the content so a repeat run drops files that the site no longer holds' {
        $plan = New-CMImportPlan -Model (Read-Fixture 'clauses')
        $plan.DeploymentTypes[0].ContentLocation = Join-Path $script:Source 'x64'
        $stage = Export-CMImportStage -Plan $plan -OutputRoot $script:Out
        $stale = Join-Path $stage.Staged[0].ContentFolder 'old-file.txt'
        Set-Content -LiteralPath $stale -Value 'x'
        [void](Export-CMImportStage -Plan $plan -OutputRoot $script:Out)
        Test-Path -LiteralPath $stale | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $stage.Staged[0].ContentFolder 'install.bat') | Should -BeTrue
    }

    It 'stops when the setup file is missing from the copied content' {
        $plan = New-CMImportPlan -Model (Read-Fixture 'clauses')
        $empty = Join-Path $script:Source 'empty'
        New-Item -ItemType Directory -Force -Path $empty | Out-Null
        Set-Content -LiteralPath (Join-Path $empty 'other.txt') -Value 'x'
        $plan.DeploymentTypes[0].ContentLocation = $empty
        { Export-CMImportStage -Plan $plan -OutputRoot $script:Out } | Should -Throw '*setup file*'
    }

    It 'skips the copy on request and leaves out a deployment type that was not converted' {
        $plan = New-CMImportPlan -Model (Read-Fixture 'unreachable')
        $stage = Export-CMImportStage -Plan $plan -OutputRoot $script:Out -SkipContent
        $stage.Staged.Count | Should -Be 1
    }
}

Describe 'Publish-IntuneWin32App with a converted manifest' {
    BeforeAll {
        function New-ImportPublishMocks {
            Mock Get-IntuneWinEncryptionInfo -ModuleName AppPackagerCommon {
                [pscustomobject]@{ SetupFile = 'install.bat'; FileName = 'w.intunewin'; UnencryptedContentSize = 10
                    EncryptionKey = 'k'; MacKey = 'm'; InitializationVector = 'i'; Mac = 'M'; ProfileIdentifier = 'p'; FileDigest = 'd'; FileDigestAlgorithm = 'SHA256' }
            }
            Mock Get-MsGraphToken -ModuleName AppPackagerCommon { 'token' }
            Mock Export-IntuneWinPayload -ModuleName AppPackagerCommon { [pscustomobject]@{ Path = (Join-Path $TestDrive 'payload.bin'); Size = 10 } }
            Mock Invoke-AzureBlobUpload -ModuleName AppPackagerCommon { }
            Mock Start-Sleep -ModuleName AppPackagerCommon { }
            Mock Invoke-GraphJson -ModuleName AppPackagerCommon {
                $script:GraphCalls.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $Body })
                if ($Method -eq 'GET' -and $Uri -match 'mobileApps\?') { return [pscustomobject]@{ value = @() } }
                if ($Method -eq 'GET') { return [pscustomobject]@{ azureStorageUri = 'https://blob.invalid/x'; uploadState = 'commitFileSuccess' } }
                return [pscustomobject]@{ id = 'app-1' }
            }
        }
    }

    BeforeEach {
        $script:GraphCalls = New-Object System.Collections.Generic.List[object]
        Set-Content -LiteralPath (Join-Path $TestDrive 'payload.bin') -Value 'x'
        Set-Content -LiteralPath (Join-Path $TestDrive 'w.intunewin') -Value 'x'
    }

    It 'sends the finished rules, return codes, minimums and the architecture set' {
        New-ImportPublishMocks
        $manifest = (New-CMImportPlan -Model (Read-Fixture 'misc-reqs')).DeploymentTypes[0].Manifest
        Publish-IntuneWin32App -TenantId 't' -ClientId 'c' -ClientSecret 's' -IntuneWinPath (Join-Path $TestDrive 'w.intunewin') -Manifest $manifest | Should -Be 'app-1'
        $create = @($script:GraphCalls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match 'mobileApps$' })[0]
        $create.Body['allowedArchitectures'] | Should -Be 'x64'
        $create.Body['minimumMemoryInMB'] | Should -Be 4096
        $create.Body['minimumNumberOfProcessors'] | Should -Be 2
        $create.Body['rules'].Count | Should -Be 1
        $fileRule = $create.Body['rules'][0]
        $fileRule['@odata.type'] | Should -Be '#microsoft.graph.win32LobAppFileSystemRule'
        $fileRule['operationType'] | Should -Be 'exists'
        $fileRule['fileOrFolderName'] | Should -Be 'app.exe'
        $fileRule['path'] | Should -Be 'C:\Program Files\ImpTestMisc'
        $fileRule['check32BitOn64System'] | Should -BeTrue
        $create.Body['returnCodes'].Count | Should -Be 5
        $create.Body['installExperience']['deviceRestartBehavior'] | Should -Be 'basedOnReturnCode'
        $create.Body['displayName'] | Should -Be 'IMPTEST Misc Reqs'
        $create.Body['notes'] | Should -Match '^AppPackager:cm:Application_.+/DeploymentType_'
    }

    It 'sends the documented property when the app ran on x86 and x64' {
        New-ImportPublishMocks
        $manifest = (New-CMImportPlan -Model (Read-Fixture 'clauses')).DeploymentTypes[0].Manifest
        $manifest.IntuneAllowedArchitectures | Should -Be 'x86,x64'
        Publish-IntuneWin32App -TenantId 't' -ClientId 'c' -ClientSecret 's' -IntuneWinPath (Join-Path $TestDrive 'w.intunewin') -Manifest $manifest | Should -Be 'app-1'
        $create = @($script:GraphCalls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match 'mobileApps$' })[0]
        $create.Body.ContainsKey('allowedArchitectures') | Should -BeFalse
        $create.Body['applicableArchitectures'] | Should -Be 'x86,x64'
        $create.Body['rules'].Count | Should -Be 4
    }

    It 'publishes a manifest read back from its JSON file' {
        New-ImportPublishMocks
        $plan = New-CMImportPlan -Model (Read-Fixture 'multi-lang')
        $path = Join-Path $TestDrive 'stage-manifest.json'
        $plan.DeploymentTypes[1].Manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $path -Encoding UTF8
        $manifest = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        Publish-IntuneWin32App -TenantId 't' -ClientId 'c' -ClientSecret 's' -IntuneWinPath (Join-Path $TestDrive 'w.intunewin') -Manifest $manifest | Should -Be 'app-1'
        $create = @($script:GraphCalls | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match 'mobileApps$' })[0]
        $create.Body['rules'].Count | Should -Be 2
        @($create.Body['rules'] | Where-Object { $_.ruleType -eq 'requirement' }).Count | Should -Be 1
        $create.Body['returnCodes'].Count | Should -Be 5
    }

    It 'gives each variant its own identity so a repeat run updates the same app' {
        New-ImportPublishMocks
        $plan = New-CMImportPlan -Model (Read-Fixture 'multi-arch')
        $tags = @($plan.DeploymentTypes | ForEach-Object { Get-IntuneIdentityTag -Manifest $_.Manifest })
        $tags.Count | Should -Be 2
        $tags[0] | Should -Not -Be $tags[1]
    }

    It 'accepts a converted manifest in the compatibility findings' {
        $manifest = (New-CMImportPlan -Model (Read-Fixture 'or-detect')).DeploymentTypes[0].Manifest
        $codes = @(Get-IntuneCompatibilityFindings -Manifest $manifest | ForEach-Object { $_.Code })
        $codes | Should -Not -Contain 'DetectionNotMappable'
        $codes | Should -Not -Contain 'ArchitectureUnresolved'
        $codes | Should -Not -Contain 'MultipleDeploymentTypes'
    }
}

Describe 'Convert-CMApplicationToIntune.ps1' {
    BeforeAll {
        $script:WrapperPath = Join-Path $PSScriptRoot '..\Convert-CMApplicationToIntune.ps1'
        $tokens = $null; $errors = $null
        $script:WrapperAst = [System.Management.Automation.Language.Parser]::ParseFile($script:WrapperPath, [ref]$tokens, [ref]$errors)
        $script:WrapperErrors = @($errors)
    }

    It 'parses' {
        $script:WrapperErrors.Count | Should -Be 0
    }

    It 'takes the documented parameters' {
        $names = @($script:WrapperAst.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        foreach ($expected in 'ApplicationName', 'SiteCode', 'OutputRoot', 'VariantStrategy', 'NativeDetectionOnly', 'AllowUntranslatedRequirements', 'SkipContent', 'Package', 'Publish', 'UseSavedSettings', 'TenantId', 'ClientId', 'ClientSecret', 'ToolPath') {
            $names | Should -Contain $expected
        }
    }

    It 'uses no loop variable that shadows a parameter' {
        $parameterNames = @($script:WrapperAst.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath.ToLowerInvariant() })
        $loops = @($script:WrapperAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.ForEachStatementAst] }, $true))
        foreach ($loop in $loops) { $parameterNames | Should -Not -Contain $loop.Variable.VariablePath.UserPath.ToLowerInvariant() }
    }

    It 'keeps the secret out of the parameter list as plain text' {
        $secret = $script:WrapperAst.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'ClientSecret' }
        $secret.StaticType.Name | Should -Be 'SecureString'
    }
}

Describe 'Convert-CMApplicationToIntune' {
    It 'plans and stages from the stored document without a site' {
        $source = Join-Path $TestDrive 'content'
        New-Item -ItemType Directory -Force -Path $source | Out-Null
        Set-Content -LiteralPath (Join-Path $source 'install.bat') -Value 'exit /b 0'
        Set-Content -LiteralPath (Join-Path $source 'uninstall.bat') -Value 'exit /b 0'
        $xml = (Get-Content -LiteralPath (Join-Path $script:FixtureRoot 'clauses.app.xml') -Raw -Encoding UTF8).Replace('\\cm01.contoso.com\ImpTest\x64\', $source + '\')
        $result = Convert-CMApplicationToIntune -ApplicationName 'IMPTEST Clauses' -OutputRoot (Join-Path $TestDrive 'run') -Xml $xml
        $result.Plan.DeploymentTypes[0].Convert | Should -BeTrue
        $result.Stage.Staged.Count | Should -Be 1
        $result.Packages.Count | Should -Be 0
        Test-Path -LiteralPath (Join-Path $result.Stage.Staged[0].ContentFolder 'install.bat') | Should -BeTrue
    }

    It 'builds the package and publishes each variant when asked' {
        $source = Join-Path $TestDrive 'content2'
        New-Item -ItemType Directory -Force -Path $source | Out-Null
        Set-Content -LiteralPath (Join-Path $source 'install.bat') -Value 'exit /b 0'
        $xml = (Get-Content -LiteralPath (Join-Path $script:FixtureRoot 'multi-arch.app.xml') -Raw -Encoding UTF8).Replace('\\cm01.contoso.com\ImpTest\x64\', $source + '\').Replace('\\cm01.contoso.com\ImpTest\x86\', $source + '\')
        $script:Published = New-Object System.Collections.Generic.List[string]
        Mock New-IntuneWinPackage -ModuleName AppPackagerCmImport { [pscustomobject]@{ IntuneWinPath = (Join-Path $TestDrive 'x.intunewin'); SizeBytes = 1; Sha256 = 'x'; ToolExitCode = 0; DurationSec = 0 } }
        Mock Publish-IntuneWin32App -ModuleName AppPackagerCmImport { $script:Published.Add([string]$Manifest.AppName); return ('id-' + $script:Published.Count) }
        $result = Convert-CMApplicationToIntune -ApplicationName 'IMPTEST Multi Arch' -OutputRoot (Join-Path $TestDrive 'run2') -Xml $xml -Publish -ToolPath (Join-Path $TestDrive 'IntuneWinAppUtil.exe') -TenantId t -ClientId c -ClientSecret s
        @($script:Published) | Should -Be @('IMPTEST Multi Arch - Multi - x64', 'IMPTEST Multi Arch - Multi - x86')
        @($result.Packages.AppId) | Should -Be @('id-1', 'id-2')
    }

    It 'needs the tenant values to publish' {
        $source = Join-Path $TestDrive 'content3'
        New-Item -ItemType Directory -Force -Path $source | Out-Null
        Set-Content -LiteralPath (Join-Path $source 'install.bat') -Value 'exit /b 0'
        $xml = (Get-Content -LiteralPath (Join-Path $script:FixtureRoot 'clauses.app.xml') -Raw -Encoding UTF8).Replace('\\cm01.contoso.com\ImpTest\x64\', $source + '\')
        Mock New-IntuneWinPackage -ModuleName AppPackagerCmImport { [pscustomobject]@{ IntuneWinPath = (Join-Path $TestDrive 'x.intunewin') } }
        { Convert-CMApplicationToIntune -ApplicationName 'IMPTEST Clauses' -OutputRoot (Join-Path $TestDrive 'run3') -Xml $xml -Publish -ToolPath 'x' } | Should -Throw '*TenantId*'
    }
}

Describe 'Edge cases found in review' {
    BeforeAll {
        function Get-FixtureXml([string]$Name) { Get-Content -LiteralPath (Join-Path $script:FixtureRoot "$Name.app.xml") -Raw -Encoding UTF8 }
        $script:RuleNamespace = 'http://schemas.microsoft.com/SystemsCenterConfigurationManager/2009/06/14/Rules'
        function New-OsRuleXml([string]$PlatformId) {
            '<Rule xmlns="{0}" id="Rule_extra" Severity="None" NonCompliantWhenSettingIsNotFound="false"><Annotation><DisplayName Text="extra"/><Description Text=""/></Annotation><OperatingSystemExpression><Operator>OneOf</Operator><Operands><RuleExpression RuleId="{1}"/></Operands></OperatingSystemExpression></Rule>' -f $script:RuleNamespace, $PlatformId
        }
        function New-DiskRuleXml {
            '<Rule xmlns="{0}" id="Rule_disk" Severity="None" NonCompliantWhenSettingIsNotFound="false"><Annotation><DisplayName Text="Disk"/><Description Text=""/></Annotation><Expression><Operator>GreaterEquals</Operator><Operands><GlobalSettingReference AuthoringScopeId="GLOBAL" LogicalName="FreeDiskSpace" DataType="Int64" SettingLogicalName="x" SettingSourceType="CIM" Method="Value"/><ConstantValue Value="5" DataType="Int64"/></Operands></Expression></Rule>' -f $script:RuleNamespace
        }
    }

    It 'doubles typographic single quotes in generated scripts' {
        $leaf = New-TestLeaf -Source File -Property Exists -Path 'C:\X' -Name ("a{0}b{1}c.txt" -f [char]0x2019, [char]0x2018)
        $text = New-IntuneExpressionDetectionScript -Tree $leaf
        Test-ParsesCleanly $text | Should -BeTrue
        $text.Contains(("a{0}{0}b{1}{1}c.txt" -f [char]0x2019, [char]0x2018)) | Should -BeTrue
    }

    It 'keeps the bitness of a script detector' {
        $xml = (Get-FixtureXml 'script-detect') -replace '(<Arg Name="RunAs32Bit" Type="Boolean">)false(</Arg></Args></DetectAction>)', '${1}true${2}'
        $model = ConvertFrom-CMApplicationDigest -Xml $xml
        $model.DeploymentTypes[0].Detection.RunAs32Bit | Should -BeTrue
        $rule = (New-CMImportPlan -Model $model).DeploymentTypes[0].Manifest.IntuneRules | Where-Object { $_['ruleType'] -eq 'detection' }
        $rule['runAs32Bit'] | Should -BeTrue
        (New-CMImportPlan -Model (Read-Fixture 'script-detect')).DeploymentTypes[0].Manifest.IntuneRules[0]['runAs32Bit'] | Should -BeFalse
    }

    It 'orders the deployment types by the application list' {
        $doc = New-Object System.Xml.XmlDocument
        $doc.LoadXml((Get-FixtureXml 'multi-arch'))
        $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $ns.AddNamespace('a', 'http://schemas.microsoft.com/SystemCenterConfigurationManager/2009/AppMgmtDigest')
        $list = $doc.SelectSingleNode('/a:AppMgmtDigest/a:Application/a:DeploymentTypes', $ns)
        $first = $list.FirstChild
        [void]$list.RemoveChild($first)
        [void]$list.AppendChild($first)
        $model = ConvertFrom-CMApplicationDigest -Xml $doc.OuterXml
        @($model.DeploymentTypes.Name) | Should -Be @('Multi - x86', 'Multi - x64')
        @($model.DeploymentTypes.Priority) | Should -Be @(1, 2)
    }

    It 'does not convert a registry key comparison that is not an existence test' {
        $xml = (Get-FixtureXml 'clauses') -replace '(<Operator>)Equals(</Operator><Operands><SettingReference [^>]*SettingSourceType="RegistryKey")', '${1}GreaterThan${2}'
        $dt = (New-CMImportPlan -Model (ConvertFrom-CMApplicationDigest -Xml $xml)).DeploymentTypes[0]
        $dt.Convert | Should -BeFalse
        $dt.Findings.Code | Should -Contain 'DetectionNotConvertible'
    }

    It 'blocks a detection expression that has no conditions' {
        $result = ConvertTo-IntuneDetectionRules -Detection ([pscustomobject]@{ Method = 'Enhanced'; Tree = (New-TestGroup 'And' @()) })
        $result.Rules.Count | Should -Be 0
        ($result.Findings | Where-Object { $_.Severity -eq 'Blocking' }).Code | Should -Be 'DetectionMissing'
    }

    It 'keeps the larger memory and processor minimum' {
        $memory = { param([long]$Bytes) [pscustomobject]@{ Kind = 'Global'; DisplayName = 'm'; Operator = 'GreaterEquals'; PlatformIds = @(); LogicalName = 'TotalPhysicalMemory'; Values = @([string]$Bytes); DataType = 'Int64'; Reason = '' } }
        $cpu = { param([int]$Count) [pscustomobject]@{ Kind = 'Global'; DisplayName = 'c'; Operator = 'GreaterEquals'; PlatformIds = @(); LogicalName = 'NumberOfProcessors'; Values = @([string]$Count); DataType = 'Int64'; Reason = '' } }
        $model = ConvertTo-IntuneRequirementModel -Requirements @((& $memory 4294967296), (& $memory 2147483648), (& $cpu 4), (& $cpu 2))
        $model.Native['MinimumMemoryInMB'] | Should -Be 4096
        $model.Native['MinimumNumberOfProcessors'] | Should -Be 4
        $reverse = ConvertTo-IntuneRequirementModel -Requirements @((& $memory 2147483648), (& $memory 4294967296))
        $reverse.Native['MinimumMemoryInMB'] | Should -Be 4096
    }

    It 'does not convert a deployment type whose operating system rules share no architecture' {
        $xml = (Get-FixtureXml 'misc-reqs') -replace '</Requirements>', ((New-OsRuleXml 'Windows/All_x86_Windows_10_higher') + '</Requirements>')
        $model = ConvertFrom-CMApplicationDigest -Xml $xml
        (ConvertTo-IntuneRequirementModel -Requirements @($model.DeploymentTypes[0].Requirements)).Architectures.Count | Should -Be 0
        $dt = (New-CMImportPlan -Model $model).DeploymentTypes[0]
        $dt.Convert | Should -BeFalse
        $dt.Findings.Code | Should -Contain 'ArchitectureConflict'
        $dt.Findings.Code | Should -Not -Contain 'ArchitectureNotConstrained'
    }

    It 'does not convert a deployment type with a requirement it cannot translate, unless allowed' {
        $xml = (Get-FixtureXml 'misc-reqs') -replace '</Requirements>', ((New-DiskRuleXml) + '</Requirements>')
        $model = ConvertFrom-CMApplicationDigest -Xml $xml
        $blocked = (New-CMImportPlan -Model $model).DeploymentTypes[0]
        $blocked.Convert | Should -BeFalse
        ($blocked.Findings | Where-Object { $_.Code -eq 'RequirementsNotTranslated' }).Severity | Should -Be 'Blocking'
        $allowed = (New-CMImportPlan -Model $model -AllowUntranslatedRequirements).DeploymentTypes[0]
        $allowed.Convert | Should -BeTrue
        ($allowed.Findings | Where-Object { $_.Code -eq 'RequirementsNotTranslated' }).Severity | Should -Be 'Review'
    }

    It 'finds a setup file in a subfolder of the content' {
        Get-ImpSetupFile -InstallCommand '.\bin\setup.exe /S' -ContentFiles @('bin\setup.exe', 'a.txt') | Should -Be 'bin\setup.exe'
        Get-ImpSetupFile -InstallCommand 'setup.exe /S' -ContentFiles @('tools\setup.exe', 'a.txt') | Should -Be 'tools\setup.exe'
        Get-ImpSetupFile -InstallCommand 'msiexec /i "pkg/app.msi" /qn' -ContentFiles @('pkg\app.msi') | Should -Be 'pkg\app.msi'
        Get-ImpSetupFile -InstallCommand 'setup.exe' -ContentFiles @('a\setup.exe', 'b\setup.exe') | Should -Be ''
    }

    It 'refuses two deployment types that map to one folder' {
        $plan = New-CMImportPlan -Model (Read-Fixture 'multi-arch')
        $plan.DeploymentTypes[0].Name = 'A/B'
        $plan.DeploymentTypes[1].Name = 'A:B'
        { Export-CMImportStage -Plan $plan -OutputRoot (Join-Path $TestDrive 'collision') -SkipContent } | Should -Throw '*same folder*'
    }

    It 'passes the staged icon to the publisher' {
        $png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='
        $source = Join-Path $TestDrive 'icon-content'
        New-Item -ItemType Directory -Force -Path $source | Out-Null
        Set-Content -LiteralPath (Join-Path $source 'install.bat') -Value 'exit /b 0'
        $xml = (Get-FixtureXml 'clauses').Replace('\\cm01.contoso.com\ImpTest\x64\', $source + '\')
        $xml = $xml.Replace('<Version>1.0</Version></Info>', ('<Version>1.0</Version><Icon Id="i"><Data>{0}</Data></Icon></Info>' -f $png))
        $script:SeenIcon = $null
        Mock New-IntuneWinPackage -ModuleName AppPackagerCmImport { [pscustomobject]@{ IntuneWinPath = (Join-Path $TestDrive 'x.intunewin') } }
        Mock Publish-IntuneWin32App -ModuleName AppPackagerCmImport { $script:SeenIcon = $IconPath; 'id-1' }
        [void](Convert-CMApplicationToIntune -ApplicationName 'IMPTEST Clauses' -OutputRoot (Join-Path $TestDrive 'icon-run') -Xml $xml -Publish -ToolPath 'x' -TenantId t -ClientId c -ClientSecret s)
        $script:SeenIcon | Should -Not -BeNullOrEmpty
        Test-Path -LiteralPath $script:SeenIcon | Should -BeTrue
    }

    It 'checks the publish credentials before it copies anything' {
        $source = Join-Path $TestDrive 'early-content'
        New-Item -ItemType Directory -Force -Path $source | Out-Null
        Set-Content -LiteralPath (Join-Path $source 'install.bat') -Value 'exit /b 0'
        $xml = (Get-FixtureXml 'clauses').Replace('\\cm01.contoso.com\ImpTest\x64\', $source + '\')
        { Convert-CMApplicationToIntune -ApplicationName 'IMPTEST Clauses' -OutputRoot (Join-Path $TestDrive 'early-run') -Xml $xml -Publish -ToolPath 'x' } | Should -Throw '*TenantId*'
        Test-Path -LiteralPath (Join-Path $TestDrive 'early-run') | Should -BeFalse
    }

    It 'refuses unsigned converted scripts when the signing policy requires signatures' {
        Mock Resolve-CommonSigningPolicy -ModuleName AppPackagerCommon { [pscustomobject]@{ RequireDetection = $true; RequireRequirements = $true } }
        $orManifest = (New-CMImportPlan -Model (Read-Fixture 'or-detect')).DeploymentTypes[0].Manifest
        @(Get-IntuneCompatibilityFindings -Manifest $orManifest | Where-Object { $_.Code -eq 'ConvertedScriptNotSigned' }).Count | Should -Be 1
        $nativeManifest = (New-CMImportPlan -Model (Read-Fixture 'clauses')).DeploymentTypes[0].Manifest
        @(Get-IntuneCompatibilityFindings -Manifest $nativeManifest | Where-Object { $_.Code -eq 'ConvertedScriptNotSigned' }).Count | Should -Be 0
        $langManifest = (New-CMImportPlan -Model (Read-Fixture 'multi-lang')).DeploymentTypes[1].Manifest
        @(Get-IntuneCompatibilityFindings -Manifest $langManifest | Where-Object { $_.Code -eq 'ConvertedScriptNotSigned' }).Count | Should -Be 1
    }

    It 'blocks finished rules that hold no detection rule' {
        $manifest = (New-CMImportPlan -Model (Read-Fixture 'multi-lang')).DeploymentTypes[1].Manifest
        $manifest.IntuneRules = @($manifest.IntuneRules | Where-Object { $_['ruleType'] -eq 'requirement' })
        (Get-IntuneCompatibilityFindings -Manifest $manifest | Where-Object { $_.Severity -eq 'Blocking' }).Code | Should -Contain 'DetectionMissing'
    }
}
