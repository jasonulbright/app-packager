
BeforeDiscovery {
    $script:PackagerCases = @(
        @{ Name = 'package-citrixworkspace-cr';         Stream = 'Current' }
        @{ Name = 'package-citrixworkspace-ltsr-x86';   Stream = 'LTSR' }
        @{ Name = 'package-citrixworkspace-ltsr-x64';   Stream = 'LTSR' }
        @{ Name = 'package-citrixworkspace-ltsr-arm64'; Stream = 'LTSR' }
    )
}

BeforeAll {
    function Get-PackagerAst {
        param([string]$Name)
        $t = $null; $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot "..\Packagers\$Name.ps1"), [ref]$t, [ref]$e)
        if ($e) { throw ($e.Message -join '; ') }
        $ast
    }
    function Get-FunctionText {
        param($Ast, [string]$Function)
        $fn = $Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Function }, $false)
        if ($fn) { $fn.Extent.Text -replace "`r`n", "`n" }
    }

    . ([scriptblock]::Create((Get-FunctionText -Ast (Get-PackagerAst 'package-citrixworkspace-cr') -Function 'ConvertTo-CwaInstallArguments')))

    function New-CwaConfig {
        param([hashtable]$Section = @{})
        $cfg = [ordered]@{}
        foreach ($k in $Section.Keys) { $cfg[$k] = [pscustomobject]$Section[$k] }
        [pscustomobject]$cfg
    }
    function Get-Args {
        param($Config, [string]$Stream = 'Current')
        (ConvertTo-CwaInstallArguments -Config $Config -Stream $Stream).Arguments
    }
}

Describe 'Citrix Workspace argument builder: baseline' {
    It 'emits /silent only for a missing config' {
        $r = ConvertTo-CwaInstallArguments -Config $null -Stream 'Current'
        $r.Arguments | Should -Be @('/silent')
        $r.Warnings | Should -BeNullOrEmpty
    }

    It 'never emits /noreboot, which the vendor install page does not document' {
        $cfg = New-CwaConfig @{ Installation = @{ CleanInstall = $true } }
        Get-Args $cfg | Should -Not -Contain '/noreboot'
    }

    It 'emits nothing beyond /silent for empty sections' {
        $cfg = New-CwaConfig @{ Installation = @{}; Plugins = @{}; UpdateAndTelemetry = @{}; StorePolicy = @{}; Store = @{}; Components = @{} }
        Get-Args $cfg | Should -Be @('/silent')
    }
}

Describe 'Citrix Workspace argument builder: store policy' {
    It 'omits <Key> when it is not set (<Value>)' -TestCases @(
        @{ Key = 'AllowAddStore'; Value = '' }
        @{ Key = 'AllowAddStore'; Value = '   ' }
        @{ Key = 'AllowAddStore'; Value = $null }
        @{ Key = 'AllowSavePwd';  Value = '' }
        @{ Key = 'AllowSavePwd';  Value = $null }
    ) {
        param($Key, $Value)
        $r = ConvertTo-CwaInstallArguments -Config (New-CwaConfig @{ StorePolicy = @{ $Key = $Value } }) -Stream 'Current'
        $r.Arguments | Should -Be @('/silent')
        $r.Warnings | Should -BeNullOrEmpty
    }

    It 'emits <Switch>=<Expected> for <Value>' -TestCases @(
        @{ Key = 'AllowAddStore'; Switch = 'ALLOWADDSTORE'; Value = 'S'; Expected = 'S' }
        @{ Key = 'AllowAddStore'; Switch = 'ALLOWADDSTORE'; Value = 'a'; Expected = 'A' }
        @{ Key = 'AllowAddStore'; Switch = 'ALLOWADDSTORE'; Value = ' N '; Expected = 'N' }
        @{ Key = 'AllowSavePwd';  Switch = 'ALLOWSAVEPWD';  Value = 'S'; Expected = 'S' }
        @{ Key = 'AllowSavePwd';  Switch = 'ALLOWSAVEPWD';  Value = 'n'; Expected = 'N' }
        @{ Key = 'AllowSavePwd';  Switch = 'ALLOWSAVEPWD';  Value = 'A'; Expected = 'A' }
    ) {
        param($Key, $Switch, $Value, $Expected)
        Get-Args (New-CwaConfig @{ StorePolicy = @{ $Key = $Value } }) | Should -Contain "$Switch=$Expected"
    }

    It 'rejects an undocumented value with a warning: <Value>' -TestCases @(
        @{ Value = 'Y' }
        @{ Value = 'Secure' }
        @{ Value = '(not set)' }
    ) {
        param($Value)
        $r = ConvertTo-CwaInstallArguments -Config (New-CwaConfig @{ StorePolicy = @{ AllowAddStore = $Value; AllowSavePwd = $Value } }) -Stream 'Current'
        $r.Arguments | Should -Be @('/silent')
        @($r.Warnings).Count | Should -Be 2
    }
}

Describe 'Citrix Workspace argument builder: installation' {
    It 'emits ENABLE_SSON only together with /includeSSON' {
        Get-Args (New-CwaConfig @{ Installation = @{ IncludeSSON = $false; EnableSSON = $true } }) | Should -Be @('/silent')
        Get-Args (New-CwaConfig @{ Installation = @{ IncludeSSON = $true; EnableSSON = $true } }) | Should -Be @('/silent', '/includeSSON', 'ENABLE_SSON=Yes')
        Get-Args (New-CwaConfig @{ Installation = @{ IncludeSSON = $true; EnableSSON = $false } }) | Should -Be @('/silent', '/includeSSON', 'ENABLE_SSON=No')
    }

    It 'uses startAppProtection, not the deprecated includeAppProtection' {
        $a = Get-Args (New-CwaConfig @{ Installation = @{ AppProtection = $true } })
        $a | Should -Contain 'startAppProtection'
        ($a -join ' ') | Should -Not -Match 'includeappprotection'
    }

    It 'emits the documented True/False values for pre-launch and self-service' {
        $a = Get-Args (New-CwaConfig @{ Installation = @{ CleanInstall = $true; SessionPreLaunch = $false; SelfServiceMode = $true } })
        $a | Should -Be @('/silent', '/CleanInstall', 'ENABLEPRELAUNCH=False', 'SELFSERVICEMODE=True')
    }
}

Describe 'Citrix Workspace argument builder: plugins' {
    It 'Current Release: Teams Y/N, Zoom and EPA only as documented opt-outs' {
        Get-Args (New-CwaConfig @{ Plugins = @{ MSTeamsPlugin = $true; ZoomPlugin = $true; EPAClient = $true } }) |
            Should -Be @('/silent', '/InstallMSTeamsPlugin=Y')
        Get-Args (New-CwaConfig @{ Plugins = @{ MSTeamsPlugin = $false; ZoomPlugin = $false; EPAClient = $false } }) |
            Should -Be @('/silent', '/InstallMSTeamsPlugin=N', 'Installzoomplugin=N', 'InstallEPAClient=N')
    }

    It 'LTSR: installs Zoom through ADDONS and ignores the Teams setting with a warning' {
        $r = ConvertTo-CwaInstallArguments -Config (New-CwaConfig @{ Plugins = @{ MSTeamsPlugin = $true; ZoomPlugin = $true; WebExPlugin = $true } }) -Stream 'LTSR'
        $r.Arguments | Should -Be @('/silent', 'ADDONS=ZoomVDIPlugin,WebexVDIPlugin')
        @($r.Warnings | Where-Object { $_ -like 'Plugins.MSTeamsPlugin*' }).Count | Should -Be 1
    }

    It 'LTSR: emits no Installzoomplugin switch' {
        Get-Args (New-CwaConfig @{ Plugins = @{ ZoomPlugin = $false } }) -Stream 'LTSR' | Should -Be @('/silent')
    }

    It 'Current Release: ADDONS carries WebEx only' {
        Get-Args (New-CwaConfig @{ Plugins = @{ ZoomPlugin = $true; WebExPlugin = $true } }) | Should -Be @('/silent', 'ADDONS=WebexVDIPlugin')
    }

    It 'emits /SkipUberAgentUpgrade after /InstallUberAgent' {
        Get-Args (New-CwaConfig @{ Plugins = @{ UberAgent = $true; UberAgentSkipUpgrade = $true } }) |
            Should -Be @('/silent', '/InstallUberAgent', '/SkipUberAgentUpgrade')
    }

    It 'warns and skips /SkipUberAgentUpgrade without /InstallUberAgent' {
        $r = ConvertTo-CwaInstallArguments -Config (New-CwaConfig @{ Plugins = @{ UberAgent = $false; UberAgentSkipUpgrade = $true } }) -Stream 'Current'
        $r.Arguments | Should -Be @('/silent')
        @($r.Warnings).Count | Should -Be 1
    }

    It 'emits /InstallSRAgent on Current Release only' {
        $cfg = New-CwaConfig @{ Plugins = @{ SessionRecording = $true } }
        Get-Args $cfg -Stream 'Current' | Should -Contain '/InstallSRAgent'
        $r = ConvertTo-CwaInstallArguments -Config $cfg -Stream 'LTSR'
        $r.Arguments | Should -Not -Contain '/InstallSRAgent'
        @($r.Warnings).Count | Should -Be 1
    }
}

Describe 'Citrix Workspace argument builder: update and telemetry' {
    It 'omits AutoUpdateCheck when not set' {
        Get-Args (New-CwaConfig @{ UpdateAndTelemetry = @{ AutoUpdateCheck = '' } }) | Should -Be @('/silent')
    }

    It 'emits AutoUpdateCheck=<Expected>' -TestCases @(
        @{ Value = 'auto';     Expected = 'auto' }
        @{ Value = 'Manual';   Expected = 'manual' }
        @{ Value = 'disabled'; Expected = 'disabled' }
    ) {
        param($Value, $Expected)
        Get-Args (New-CwaConfig @{ UpdateAndTelemetry = @{ AutoUpdateCheck = $Value } }) | Should -Be @('/silent', "AutoUpdateCheck=$Expected")
    }

    It 'rejects an undocumented AutoUpdateCheck value' {
        $r = ConvertTo-CwaInstallArguments -Config (New-CwaConfig @{ UpdateAndTelemetry = @{ AutoUpdateCheck = 'off' } }) -Stream 'Current'
        $r.Arguments | Should -Be @('/silent')
        @($r.Warnings).Count | Should -Be 1
    }

    It 'emits EnableCEIP and EnableTracing in the vendor spelling' {
        Get-Args (New-CwaConfig @{ UpdateAndTelemetry = @{ EnableCEIP = $false; EnableTracing = $false } }) |
            Should -Be @('/silent', 'EnableCEIP=False', 'EnableTracing=false')
    }
}

Describe 'Citrix Workspace argument builder: STORE0' {
    It 'passes a StoreFront discovery URL as typed' {
        Get-Args (New-CwaConfig @{ Store = @{ Name = 'HR'; Url = 'https://sf.example.com/Citrix/Store/discovery' } }) |
            Should -Be @('/silent', 'STORE0=HR;https://sf.example.com/Citrix/Store/discovery;On;HR')
    }

    It 'passes a Gateway URL with a store fragment as typed' {
        Get-Args (New-CwaConfig @{ Store = @{ Name = 'HR'; Url = 'https://ag.example.com#Store' } }) |
            Should -Contain 'STORE0=HR;https://ag.example.com#Store;On;HR'
    }

    It 'quotes the value when the store name has a space' {
        Get-Args (New-CwaConfig @{ Store = @{ Name = 'HR Store'; Url = 'https://sf.example.com/Citrix/Store/discovery' } }) |
            Should -Contain 'STORE0="HR Store;https://sf.example.com/Citrix/Store/discovery;On;HR Store"'
    }

    It 'adds no store and no warning when name and URL are blank' {
        $r = ConvertTo-CwaInstallArguments -Config (New-CwaConfig @{ Store = @{ Name = ''; Url = '' } }) -Stream 'Current'
        $r.Arguments | Should -Be @('/silent')
        $r.Warnings | Should -BeNullOrEmpty
    }

    It 'rejects <Case> with a warning' -TestCases @(
        @{ Case = 'a URL without a scheme'; Name = 'Workspace'; Url = 'cloud.example.com' }
        @{ Case = 'a non-http scheme';      Name = 'Workspace'; Url = 'ftp://cloud.example.com' }
        @{ Case = 'a name with a semicolon'; Name = 'A;B';     Url = 'https://sf.example.com/discovery' }
        @{ Case = 'a name without a URL';   Name = 'Workspace'; Url = '' }
        @{ Case = 'a URL without a name';   Name = '';          Url = 'https://sf.example.com/discovery' }
    ) {
        param($Case, $Name, $Url)
        $r = ConvertTo-CwaInstallArguments -Config (New-CwaConfig @{ Store = @{ Name = $Name; Url = $Url } }) -Stream 'Current'
        $r.Arguments | Should -Be @('/silent')
        @($r.Warnings).Count | Should -Be 1
    }
}

Describe 'Citrix Workspace argument builder: ADDLOCAL' {
    It 'omits ADDLOCAL unless Customize is set' {
        Get-Args (New-CwaConfig @{ Components = @{ Customize = $false; ReceiverInside = $true } }) | Should -Be @('/silent')
    }

    It 'lists the selected components in vendor spelling' {
        Get-Args (New-CwaConfig @{ Components = @{ Customize = $true; ReceiverInside = $true; ICA_Client = $true; AM = $true; SelfService = $true; USB = $false } }) |
            Should -Be @('/silent', 'ADDLOCAL=ReceiverInside,ICA_Client,AM,SelfService')
    }
}

Describe 'Citrix Workspace packager <Name>' -ForEach $PackagerCases {
    BeforeAll {
        $script:Ast = Get-PackagerAst $Name
        $script:ReferenceText = Get-FunctionText -Ast (Get-PackagerAst 'package-citrixworkspace-cr') -Function 'ConvertTo-CwaInstallArguments'
    }

    It 'carries the same argument builder as the Current Release packager' {
        Get-FunctionText -Ast $Ast -Function 'ConvertTo-CwaInstallArguments' | Should -BeExactly $ReferenceText
    }

    It 'passes its own stream to the builder' {
        $node = $Ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$CwaStream' }, $false)
        $node.Right.Extent.Text | Should -Be "'$Stream'"
    }

    It 'no longer defines the removed Get-CwaBoolText helper' {
        Get-FunctionText -Ast $Ast -Function 'Get-CwaBoolText' | Should -BeNullOrEmpty
    }
}
