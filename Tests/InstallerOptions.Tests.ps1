#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The installer switches and properties that the packagers build from the Options settings.

.DESCRIPTION
    Each switch has a vendor source in docs\reference\installer-options-audit.md. The cases hold the
    syntax of every switch that the audit fixed or added: SSMS, Adobe Acrobat Reader, Beyond Compare 5
    and DBeaver. The Citrix Workspace App and Microsoft 365 switches have their own test files.
#>

BeforeAll {
    # The functions load into this scope, so the loop does the dot-sourcing and no helper function does.
    $script:PackagerAst = @{}
    foreach ($spec in @(
            @{ Packager = 'package-ssms'; Name = @('Get-SsmsOptionList', 'Get-SsmsInstallArgs', 'Get-SsmsUninstallArgs', 'Get-SsmsInstallPath') }
            @{ Packager = 'package-adobereader'; Name = @('Get-AdobePropertyCommandLine', 'Get-AdobeMuiCommandLine') }
            @{ Packager = 'package-beyondcompare5'; Name = @('Get-BeyondCompare5InstallSwitch') })) {
        $t = $null; $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot ('..\Packagers\{0}.ps1' -f $spec.Packager)), [ref]$t, [ref]$e)
        if ($e) { throw ($e.Message -join '; ') }
        $script:PackagerAst[$spec.Packager] = $ast
        foreach ($n in $spec.Name) {
            $fn = $ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $false)
            if (-not $fn) { throw ('{0} no longer defines {1}.' -f $spec.Packager, $n) }
            . ([scriptblock]::Create($fn.Extent.Text))
        }
    }
    $script:DefaultInstallPath = 'C:\Program Files\Microsoft SQL Server Management Studio 22\Release'
    $script:BaseInstallSwitches = @('/VERYSILENT', '/NORESTART', '/ALLUSERS', '/DISABLEUPDATES', '/SUPPRESSMSGBOXES')

    function New-SsmsOption {
        param([hashtable]$Override = @{})
        $o = [ordered]@{ UIMode = 'Quiet'; DownloadThenInstall = $false; NoUpdateInstaller = $false; IncludeRecommended = $false; IncludeOptional = $false; RemoveOos = $false; ForceClose = $false; InstallPath = ''; Workloads = @(); Languages = @(); NoCache = $false }
        foreach ($k in $Override.Keys) { $o[$k] = $Override[$k] }
        [pscustomobject]$o
    }
}

Describe 'SSMS 22 install arguments' {
    It 'starts with --quiet and --norestart' {
        Get-SsmsInstallArgs -Options (New-SsmsOption) | Should -Be @('--quiet', '--norestart')
    }

    It 'uses --passive in place of --quiet for the Passive mode' {
        Get-SsmsInstallArgs -Options (New-SsmsOption @{ UIMode = 'Passive' }) | Should -Be @('--passive', '--norestart')
    }

    It 'adds one --add for each workload, in the order of the list' {
        $o = New-SsmsOption @{ Workloads = @('Microsoft.SqlServer.Workload.SSMS.AI', 'Microsoft.SqlServer.Workload.SSMS.BI') }
        Get-SsmsInstallArgs -Options $o | Should -Be @('--quiet', '--norestart', '--add', 'Microsoft.SqlServer.Workload.SSMS.AI', '--add', 'Microsoft.SqlServer.Workload.SSMS.BI')
    }

    It 'emits --includeRecommended and --includeOptional only with a workload, because they act on the workloads of --add' {
        $without = New-SsmsOption @{ IncludeRecommended = $true; IncludeOptional = $true }
        Get-SsmsInstallArgs -Options $without | Should -Be @('--quiet', '--norestart')
        $with = New-SsmsOption @{ IncludeRecommended = $true; IncludeOptional = $true; Workloads = @('Microsoft.SqlServer.Workload.SSMS.DevOps') }
        Get-SsmsInstallArgs -Options $with | Should -Be @('--quiet', '--norestart', '--add', 'Microsoft.SqlServer.Workload.SSMS.DevOps', '--includeRecommended', '--includeOptional')
    }

    It 'drops a workload name that is not a SSMS workload ID' {
        $o = New-SsmsOption @{ Workloads = @('x; calc', 'Microsoft.SqlServer.Workload.SSMS.AI') }
        Get-SsmsInstallArgs -Options $o | Should -Be @('--quiet', '--norestart', '--add', 'Microsoft.SqlServer.Workload.SSMS.AI')
    }

    It 'adds --addProductLang for each language and --nocache' {
        $o = New-SsmsOption @{ Languages = @('de-de', 'fr-fr'); NoCache = $true }
        Get-SsmsInstallArgs -Options $o | Should -Be @('--quiet', '--norestart', '--addProductLang', 'de-de', '--addProductLang', 'fr-fr', '--nocache')
    }

    It 'emits --noUpdateInstaller with --quiet and not with --passive' {
        Get-SsmsInstallArgs -Options (New-SsmsOption @{ NoUpdateInstaller = $true }) | Should -Contain '--noUpdateInstaller'
        Get-SsmsInstallArgs -Options (New-SsmsOption @{ NoUpdateInstaller = $true; UIMode = 'Passive' }) | Should -Not -Contain '--noUpdateInstaller'
    }

    It 'keeps --downloadThenInstall, --removeOos true, --force and --installPath' {
        $o = New-SsmsOption @{ DownloadThenInstall = $true; RemoveOos = $true; ForceClose = $true; InstallPath = 'D:\SSMS' }
        Get-SsmsInstallArgs -Options $o | Should -Be @('--quiet', '--norestart', '--downloadThenInstall', '--removeOos', 'true', '--force', '--installPath', 'D:\SSMS')
    }

    It 'builds the uninstall from the same install path' {
        Get-SsmsUninstallArgs -Options (New-SsmsOption @{ InstallPath = 'D:\SSMS' }) | Should -Be @('uninstall', '--installPath', 'D:\SSMS', '--quiet', '--norestart')
    }

    It 'reads the new options in Get-SsmsInstallOptions' {
        $text = ($script:PackagerAst['package-ssms'].Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-SsmsInstallOptions' }, $false)).Extent.Text
        $text | Should -Match 'Workloads'
        $text | Should -Match 'NoCache'
    }
}

Describe 'Adobe Acrobat Reader property options' {
    It 'emits nothing for the defaults' {
        Get-AdobePropertyCommandLine -Options ([pscustomobject]@{ AcceptEula = $false; DisableDesktopShortcut = $false; DisableChromeExtension = $false; DisableBrowserIntegration = $false; UpdateMode = '' }) | Should -Be ''
    }

    It 'emits each documented property with its documented value' {
        $o = [pscustomobject]@{ AcceptEula = $true; DisableDesktopShortcut = $true; DisableChromeExtension = $true; DisableBrowserIntegration = $true; UpdateMode = '3' }
        Get-AdobePropertyCommandLine -Options $o | Should -Be 'EULA_ACCEPT=YES DISABLEDESKTOPSHORTCUT=1 ENABLE_CHROMEEXT=0 DISABLE_BROWSER_INTEGRATION=YES UPDATE_MODE=3'
    }

    It 'accepts only the UPDATE_MODE values that the Adobe page lists' -ForEach @(
        @{ Mode = '0'; Expected = 'UPDATE_MODE=0' }
        @{ Mode = '2'; Expected = 'UPDATE_MODE=2' }
        @{ Mode = '4'; Expected = 'UPDATE_MODE=4' }
        @{ Mode = '1'; Expected = '' }
        @{ Mode = '9'; Expected = '' }
    ) {
        Get-AdobePropertyCommandLine -Options ([pscustomobject]@{ AcceptEula = $false; DisableDesktopShortcut = $false; DisableChromeExtension = $false; DisableBrowserIntegration = $false; UpdateMode = $Mode }) | Should -Be $Expected
    }

    It 'keeps the language properties of the MUI edition' {
        Get-AdobeMuiCommandLine -Languages @('de_DE', 'en_US') | Should -Be 'LANG_LIST="de_DE,en_US" SUPPRESSLANGSELECTION=1'
        Get-AdobeMuiCommandLine -Languages @('All') | Should -Be 'LANG_LIST="ALL" SUPPRESSLANGSELECTION=1'
    }
}

Describe 'Beyond Compare 5 install switches' {
    It 'keeps the base switches without options' {
        Get-BeyondCompare5InstallSwitch -Preferences $null | Should -Be $script:BaseInstallSwitches
        Get-BeyondCompare5InstallSwitch -Preferences ([pscustomobject]@{ BeyondCompareOptions = [pscustomobject]@{ DisableRemoteProfiles = $false; DisableSavedPasswords = $false } }) | Should -Be $script:BaseInstallSwitches
    }

    It 'appends the policy switches that are selected' {
        $prefs = [pscustomobject]@{ BeyondCompareOptions = [pscustomobject]@{ DisableRemoteProfiles = $true; DisableSavedPasswords = $true } }
        Get-BeyondCompare5InstallSwitch -Preferences $prefs | Should -Be ($script:BaseInstallSwitches + @('/DisableRemoteProfiles', '/DisableSavedPasswords'))
    }
}

Describe 'DBeaver uninstall arguments' {
    It 'passes /allusers with /S to the uninstaller of a System install, because /S needs a scope switch' {
        $text = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\Packagers\package-dbeaver.ps1') -Raw
        $text | Should -Match '\$uninstallArgs = "/allusers /S"'
        $text | Should -Not -Match '\$uninstallArgs = "/S"'
    }

    It 'documents the environment variable, not a dbeaver.ini line, for the AI setting' {
        $text = Get-Content -LiteralPath (Join-Path $PSScriptRoot '..\Packagers\package-dbeaver.ps1') -Raw
        $text | Should -Match 'DBEAVER_AI_DISABLED'
        $text | Should -Not -Match 'Appends -Dai\.disabled=true'
    }
}
