BeforeAll {
    $t = $null; $e = $null
    $script:GuiPath = Join-Path $PSScriptRoot '..\start-apppackager.ps1'
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:GuiPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Get-SidebarSystemState', 'Get-CmImportRunSpec', 'Get-CmImportOutcomeText', 'Get-CmImportChildEnvironment', 'Test-PipelineBusy') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $script:Xaml = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\MainWindow.xaml') -Raw

    function New-ConvertPrefs {
        param([bool]$ConfigMgr = $true, [bool]$ConsoleFound = $true, [string]$SiteCode = 'MCM')
        [pscustomobject]@{
            SiteCode      = $SiteCode
            Systems       = [pscustomobject]@{ ConfigMgr = $ConfigMgr; Intune = $false; Wsus = $false }
            DetectedTools = [pscustomobject]@{ ConfigMgrConsole = [pscustomobject]@{ Found = $ConsoleFound } }
            Intune        = [pscustomobject]@{ TenantId = ''; ClientId = ''; ClientSecretProtected = '' }
            Wsus          = [pscustomobject]@{ ServerName = '' }
            DiscoverySqlServer = ''; DiscoverySqlDatabase = ''
        }
    }
}

Describe 'Convert ConfigMgr App to Intune sidebar button' {
    It 'is in the sidebar XAML, with a wrapped label' {
        $script:Xaml | Should -Match 'x:Name="btnConvertCm"'
        $script:Xaml | Should -Match 'Convert ConfigMgr App to Intune'
    }

    It 'is wired into the window, the theme list, the enable switch and the system state' {
        $gui = Get-Content -LiteralPath $script:GuiPath -Raw
        $gui | Should -Match "\`$btnConvertCm\s*=\s*\`$window\.FindName\('btnConvertCm'\)"
        $gui | Should -Match '\$script:WorkflowButtons = @\([^)]*\$btnConvertCm'
        $gui | Should -Match '\$btnConvertCm\.IsEnabled\s*=\s*\$Enabled'
        $gui | Should -Match "@\(\`$btnConvertCm, 'ConvertCm'\)"
        $gui | Should -Match '\$btnConvertCm\.Add_Click'
    }

    It 'is enabled when ConfigMgr is in use and the console is found, even with no saved site code' {
        $s = Get-SidebarSystemState -Prefs (New-ConvertPrefs -SiteCode '')
        $s.ConvertCmEnabled | Should -BeTrue
        $s.ConvertCmToolTip | Should -Match 'Convert an existing ConfigMgr application'
    }

    It 'is disabled with a reason when the console is missing' {
        $s = Get-SidebarSystemState -Prefs (New-ConvertPrefs -ConsoleFound $false)
        $s.ConvertCmEnabled | Should -BeFalse
        $s.ConvertCmToolTip | Should -BeLike 'Convert ConfigMgr App to Intune: Needs the ConfigMgr console*'
    }

    It 'is disabled with a reason when ConfigMgr is not in use' {
        $s = Get-SidebarSystemState -Prefs (New-ConvertPrefs -ConfigMgr $false)
        $s.ConvertCmEnabled | Should -BeFalse
        $s.ConvertCmToolTip | Should -BeLike 'Convert ConfigMgr App to Intune: ConfigMgr is not in use*'
    }
}

Describe 'Get-CmImportRunSpec' {
    BeforeAll {
        function Get-Spec {
            param([hashtable]$With = @{})
            $p = @{ ApplicationName = 'Contoso Viewer'; SiteCode = 'mcm'; OutputRoot = 'D:\Converted'; ScriptPath = 'C:\AP\Convert-CMApplicationToIntune.ps1' }
            foreach ($key in $With.Keys) { $p[$key] = $With[$key] }
            Get-CmImportRunSpec @p
        }
    }

    It 'builds the command line of the wrapper script for a conversion' {
        $spec = Get-Spec
        $spec.Ok | Should -BeTrue
        $spec.Arguments | Should -Contain '-File'
        $spec.Arguments[($spec.Arguments.IndexOf('-File') + 1)] | Should -BeExactly 'C:\AP\Convert-CMApplicationToIntune.ps1'
        $spec.Arguments[($spec.Arguments.IndexOf('-ApplicationName') + 1)] | Should -BeExactly 'Contoso Viewer'
        $spec.Arguments[($spec.Arguments.IndexOf('-SiteCode') + 1)] | Should -BeExactly 'mcm'
        $spec.Arguments[($spec.Arguments.IndexOf('-OutputRoot') + 1)] | Should -BeExactly 'D:\Converted'
        $spec.Arguments[($spec.Arguments.IndexOf('-VariantStrategy') + 1)] | Should -BeExactly 'PerDeploymentType'
        $spec.Arguments | Should -Not -Contain '-SkipContent'
        $spec.Arguments | Should -Not -Contain '-Package'
        $spec.Arguments | Should -Not -Contain '-Publish'
        $spec.Summary | Should -BeLike "Conversion of 'Contoso Viewer' on site MCM*"
    }

    It 'adds -SkipContent for a preview and ignores package and publish' {
        $spec = Get-Spec @{ Mode = 'Preview'; Package = $true; Publish = $true; IntuneReady = $true }
        $spec.Ok | Should -BeTrue
        $spec.Arguments | Should -Contain '-SkipContent'
        $spec.Arguments | Should -Not -Contain '-Package'
        $spec.Arguments | Should -Not -Contain '-Publish'
        $spec.Summary | Should -BeLike 'Preview of*no content copied*'
    }

    It 'passes the strategy and the two switches' {
        $spec = Get-Spec @{ VariantStrategy = 'AssignByGroup'; NativeDetectionOnly = $true; AllowUntranslatedRequirements = $true }
        $spec.Arguments[($spec.Arguments.IndexOf('-VariantStrategy') + 1)] | Should -BeExactly 'AssignByGroup'
        $spec.Arguments | Should -Contain '-NativeDetectionOnly'
        $spec.Arguments | Should -Contain '-AllowUntranslatedRequirements'
    }

    It 'passes -Package with the tool path only when there is one' {
        $with = Get-Spec @{ Package = $true; ToolPath = 'C:\Tools\IntuneWinAppUtil.exe' }
        $with.Arguments | Should -Contain '-Package'
        $with.Arguments[($with.Arguments.IndexOf('-ToolPath') + 1)] | Should -BeExactly 'C:\Tools\IntuneWinAppUtil.exe'
        $without = Get-Spec @{ Package = $true }
        $without.Arguments | Should -Contain '-Package'
        $without.Arguments | Should -Not -Contain '-ToolPath'
    }

    It 'makes publish imply package and read the saved settings' {
        $spec = Get-Spec @{ Publish = $true; IntuneReady = $true }
        $spec.Ok | Should -BeTrue
        $spec.Arguments | Should -Contain '-Package'
        $spec.Arguments | Should -Contain '-Publish'
        $spec.Arguments | Should -Contain '-UseSavedSettings'
        $spec.Summary | Should -BeLike '*publish to Intune*'
    }

    It 'refuses to publish when the Intune settings are not saved' {
        $spec = Get-Spec @{ Publish = $true; IntuneReady = $false }
        $spec.Ok | Should -BeFalse
        $spec.Error | Should -BeLike 'Publish needs the Intune tenant*'
        @($spec.Arguments).Count | Should -Be 0
    }

    It 'does not pass the client secret on the command line' {
        $spec = Get-Spec @{ Publish = $true; IntuneReady = $true }
        ($spec.Arguments -join ' ') | Should -Not -Match '-ClientSecret|-TenantId|-ClientId'
    }

    It 'refuses an empty application name' {
        $spec = Get-Spec @{ ApplicationName = '   ' }
        $spec.Ok | Should -BeFalse
        $spec.Error | Should -BeLike 'Enter the exact name*'
    }

    It 'refuses a control character in the application name' {
        (Get-Spec @{ ApplicationName = "Bad`nName" }).Ok | Should -BeFalse
    }

    It 'refuses a site code that is not three letters or digits' {
        foreach ($bad in '', 'MC', 'MCMX', 'M-M') {
            $spec = Get-Spec @{ SiteCode = $bad }
            $spec.Ok | Should -BeFalse -Because "site code '$bad'"
            $spec.Error | Should -BeLike 'The site code is three*'
        }
    }

    It 'refuses a missing or relative output folder' {
        (Get-Spec @{ OutputRoot = '' }).Error | Should -BeLike 'Enter the output folder*'
        (Get-Spec @{ OutputRoot = 'Converted\Apps' }).Error | Should -BeLike 'The output folder must be a full path*'
    }

    It 'trims the values' {
        $spec = Get-CmImportRunSpec -ApplicationName '  Contoso Viewer ' -SiteCode ' MCM ' -OutputRoot ' D:\Converted ' -ScriptPath 'C:\AP\x.ps1'
        $spec.Arguments[($spec.Arguments.IndexOf('-ApplicationName') + 1)] | Should -BeExactly 'Contoso Viewer'
        $spec.Arguments[($spec.Arguments.IndexOf('-SiteCode') + 1)] | Should -BeExactly 'MCM'
        $spec.Arguments[($spec.Arguments.IndexOf('-OutputRoot') + 1)] | Should -BeExactly 'D:\Converted'
    }

    It 'keeps an application name with spaces, quotes and an apostrophe as one argument' {
        $spec = Get-Spec @{ ApplicationName = "Bob's ""Big"" App" }
        $spec.Ok | Should -BeTrue
        $spec.Arguments[($spec.Arguments.IndexOf('-ApplicationName') + 1)] | Should -BeExactly "Bob's ""Big"" App"
    }
}

Describe 'Get-CmImportOutcomeText' {
    It 'reports a finished preview and a finished conversion' {
        Get-CmImportOutcomeText -ExitCode 0 -Mode Preview -OutputRoot 'D:\Out' | Should -BeExactly 'Preview finished. The report is in D:\Out.'
        Get-CmImportOutcomeText -ExitCode 0 -Mode Convert -OutputRoot 'D:\Out' | Should -BeExactly 'Conversion finished. The output is in D:\Out.'
    }

    It 'shows the first line of the error text on failure' {
        $text = Get-CmImportOutcomeText -ExitCode 1 -StdErr "`r`n  Could not connect to ConfigMgr site 'ABC'.`r`nat line 5"
        $text | Should -BeExactly "Convert ConfigMgr App to Intune failed: Could not connect to ConfigMgr site 'ABC'."
    }

    It 'points to the log when there is no error text' {
        Get-CmImportOutcomeText -ExitCode 3 -StdErr '' | Should -BeExactly 'Convert ConfigMgr App to Intune failed with exit code 3. The log has the details.'
    }
}

Describe 'Get-CmImportRunSpec output folder characters' {
    It 'refuses a character that IsPathRooted throws on, with a note instead of an exception' {
        foreach ($bad in 'D:\Out"put', 'D:\Out<put', 'D:\Out>put', 'D:\Out|put', "D:\Out`tput") {
            { Get-CmImportRunSpec -ApplicationName 'App' -SiteCode 'MCM' -OutputRoot $bad -ScriptPath 'C:\AP\x.ps1' } | Should -Not -Throw
            $spec = Get-CmImportRunSpec -ApplicationName 'App' -SiteCode 'MCM' -OutputRoot $bad -ScriptPath 'C:\AP\x.ps1'
            $spec.Ok | Should -BeFalse -Because "output folder '$bad'"
            $spec.Error | Should -BeLike 'The output folder holds a character*'
        }
    }
}

Describe 'Get-CmImportChildEnvironment' {
    It 'points SMS_ADMIN_UI_PATH at the detected console folder when the parent has none' {
        $env1 = Get-CmImportChildEnvironment -ModulePath 'E:\Apps\CM Console\bin\ConfigurationManager.psd1'
        $env1['SMS_ADMIN_UI_PATH'] | Should -BeExactly 'E:\Apps\CM Console\bin\i386'
        (Split-Path -Parent $env1['SMS_ADMIN_UI_PATH']) | Should -BeExactly 'E:\Apps\CM Console\bin'
    }

    It 'leaves SMS_ADMIN_UI_PATH alone when the parent already has it' {
        $env1 = Get-CmImportChildEnvironment -ModulePath 'E:\Apps\bin\ConfigurationManager.psd1' -ParentAdminUiPath 'C:\CM\AdminConsole\bin\i386'
        $env1.Contains('SMS_ADMIN_UI_PATH') | Should -BeFalse
    }

    It 'passes the provider from the preferences unless the parent has one' {
        (Get-CmImportChildEnvironment -ProviderMachineName ' cm01.contoso.com ')['APP_PACKAGER_CM_PROVIDER'] | Should -BeExactly 'cm01.contoso.com'
        (Get-CmImportChildEnvironment -ProviderMachineName 'cm01' -ParentProvider 'cm02').Contains('APP_PACKAGER_CM_PROVIDER') | Should -BeFalse
    }

    It 'returns nothing when there is nothing to pass' {
        @((Get-CmImportChildEnvironment).Keys).Count | Should -Be 0
    }
}

Describe 'The conversion run blocks other work' {
    It 'makes Test-PipelineBusy true while the conversion runs' {
        $script:BgState = $null
        $script:CmImportRunning = $false
        Test-PipelineBusy | Should -BeFalse
        $script:CmImportRunning = $true
        Test-PipelineBusy | Should -BeTrue
        $script:CmImportRunning = $false
    }

    It 'sets the flag around the child process and clears it in finally' {
        $gui = Get-Content -LiteralPath $script:GuiPath -Raw
        $run = [regex]::Match($gui, '(?s)function Start-CmImportRun \{.*?\r?\n\}\r?\n').Value
        $run | Should -Match '\$script:CmImportRunning = \$true\s+try \{'
        $run | Should -Match '(?s)finally \{\s+\$script:CmImportRunning = \$false'
        $run | Should -Match '(?s)try \{\s+\$logFolder'
    }

    It 'stops the drop intake and the window close while the conversion runs' {
        $gui = Get-Content -LiteralPath $script:GuiPath -Raw
        $gui | Should -Match '\$script:CmImportRunning\) \{\s+\[void\]\(Show-ThemedMessage -Owner \$window -Title ''Pipeline Running'''
        $gui | Should -Match '(?s)\$window\.Add_Closing\(\{\s+param\(\$closingSender, \$closingArgs\).*?if \(\$script:CmImportRunning\) \{\s+\$closingArgs\.Cancel = \$true'
    }
}