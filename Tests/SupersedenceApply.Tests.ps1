#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The Apply Supersedence action: the engine entry point, the sidebar availability and the wiring of the button.

.DESCRIPTION
    The cases are the cases of the tests of App Packager Desktop (LocalFileGridStageTests, ApplySupersedence_RefusesWithoutAStage_...)
    and the cases of the step that the action runs (SupersedenceStepTests, the action without a publish). Every refusal comes before
    the first call to the site. The step runs against mocked ConfigMgr cmdlets: it names applications and deployment types and not
    identifiers, reads the relations before it writes, and a repeat adds only the missing ones.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:Marker = 'APD-Line:catalog:vendor-app|default|x64'
    $script:Site = 'MCM'

    function New-TestManifestFile {
        # Writes a stage manifest of the old schema (no file hashes) and returns its path.
        param([string]$Name = 'Vendor App 2.0', [string]$Version = '2.0.0', [string]$ApplicationId = 'catalog:vendor-app', [string]$TitleMode = '')
        $folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $data = [ordered]@{
            SchemaVersion = 2; AppName = $Name; SoftwareVersion = $Version; Architecture = 'x64'; ProfileId = 'default'
            Detection = [ordered]@{ Type = 'File'; FilePath = '%ProgramFiles%\Vendor'; FileName = 'app.exe'; ExpectedValue = $Version; Operator = 'GreaterEquals'; PropertyType = 'Version' }
        }
        if ($ApplicationId) { $data['ApplicationId'] = $ApplicationId }
        if ($TitleMode) { $data['TitleMode'] = $TitleMode }
        $path = Join-Path $folder 'stage-manifest.json'
        ($data | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $path -Encoding UTF8
        return $path
    }
    function New-TestSettings {
        param([string]$Mode = 'Replace', [string]$Rule = 'Minimum')
        $s = New-SupersedenceSettings
        $s.Mode = $Mode
        $s.ConfigMgrVersionRule = $Rule
        return $s
    }

    # The ConfigMgr cmdlets exist only with the console, so the module scope gets stubs that the cases mock.
    InModuleScope AppPackagerCommon {
        function script:Get-CMApplication { param($Name, [switch]$DisableWildcardHandling, $ErrorAction) }
        function script:Get-CMDeploymentType { param($ApplicationName, $DeploymentTypeName, [switch]$DisableWildcardHandling, $ErrorAction) }
        function script:Get-CMDeploymentTypeSupersedence { param($InputObject, $ErrorAction) }
        function script:Set-CMApplicationSupersedence { param($Name, $CurrentDeploymentTypeName, $SupersededApplicationName, $OldDeploymentTypeName, $IsUninstall, [switch]$DisableWildcardHandling, $ErrorAction) }
    }
}

AfterAll {
    InModuleScope AppPackagerCommon {
        foreach ($n in 'Get-CMApplication', 'Get-CMDeploymentType', 'Get-CMDeploymentTypeSupersedence', 'Set-CMApplicationSupersedence') { Remove-Item -LiteralPath "function:script:$n" -ErrorAction SilentlyContinue }
    }
}

Describe 'Invoke-SupersedenceApply' {
    BeforeEach {
        Mock Write-Log -ModuleName AppPackagerCommon {}
        Mock Connect-CMSite -ModuleName AppPackagerCommon { $true }
        Mock Get-CMApplication -ModuleName AppPackagerCommon {
            @(
                [pscustomobject]@{ CI_ID = 1; LocalizedDisplayName = 'Vendor App 1.0'; SoftwareVersion = '1.0.0'; IsExpired = $false; SDMPackageXML = '<k>APD-Line:catalog:vendor-app|default|x64</k>' }
                [pscustomobject]@{ CI_ID = 2; LocalizedDisplayName = 'Vendor App 2.0'; SoftwareVersion = '2.0.0'; IsExpired = $false; SDMPackageXML = '<none/>' }
            )
        }
        Mock Get-CMDeploymentType -ModuleName AppPackagerCommon { [pscustomobject]@{ CI_ID = 99; LocalizedDisplayName = ($ApplicationName + ' - x64') } }
        Mock Get-CMDeploymentTypeSupersedence -ModuleName AppPackagerCommon { @() }
        Mock Set-CMApplicationSupersedence -ModuleName AppPackagerCommon {}
    }

    Context 'the refusals that come before any call to the site' {
        It 'refuses a row without a stage manifest, then a staged row without a site code, and contacts no site' {
            $none = Invoke-SupersedenceApply -ManifestPath (Join-Path $TestDrive 'missing\stage-manifest.json') -SiteCode $script:Site -Settings (New-TestSettings)
            $none.Ok | Should -BeFalse
            $none.Message | Should -BeLike 'No stage manifest was found*'
            $none.RowStatus | Should -Be 'Supersedence not applied'
            $blank = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode '' -Settings (New-TestSettings)
            $blank.Ok | Should -BeFalse
            $blank.Message | Should -Match 'site code'
            Should -Invoke Connect-CMSite -ModuleName AppPackagerCommon -Times 0 -Exactly
            Should -Invoke Get-CMApplication -ModuleName AppPackagerCommon -Times 0 -Exactly
        }

        It 'refuses an empty path' {
            (Invoke-SupersedenceApply -ManifestPath '' -SiteCode $script:Site -Settings (New-TestSettings)).Message | Should -BeLike 'No stage manifest was found*'
        }

        It 'refuses a stage manifest that cannot be read and names the file' {
            $path = Join-Path $TestDrive 'broken-manifest.json'
            Set-Content -LiteralPath $path -Value '{ not json' -Encoding UTF8
            $result = Invoke-SupersedenceApply -ManifestPath $path -SiteCode $script:Site -Settings (New-TestSettings)
            $result.Ok | Should -BeFalse
            $result.Message | Should -BeLike 'The stage manifest * could not be read:*'
            Should -Invoke Connect-CMSite -ModuleName AppPackagerCommon -Times 0 -Exactly
        }

        It 'refuses with the mode Off and says where to turn it on' {
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings (New-TestSettings -Mode 'Off')
            $result.Ok | Should -BeFalse
            $result.Message | Should -Be 'Supersedence is off for this title. Turn it on in Options or in the override of the title.'
            Should -Invoke Connect-CMSite -ModuleName AppPackagerCommon -Times 0 -Exactly
        }

        It 'turns the mode on for one title with its override and not for another' {
            $settings = New-TestSettings -Mode 'Off'
            $settings.TitleOverrides['catalog:vendor-app'] = [pscustomobject]@{ Mode = 'Replace' }
            (Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings $settings).Ok | Should -BeTrue
            (Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile -ApplicationId 'catalog:other') -SiteCode $script:Site -Settings $settings).Message | Should -Match 'is off'
        }

        It 'refuses the blocked combination of Replace and uninstall with Minimum before any site call' {
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings (New-TestSettings -Mode 'ReplaceAndUninstall' -Rule 'Minimum')
            $result.Ok | Should -BeFalse
            $result.Message | Should -Match 'Replace and uninstall needs the version rule Exact'
            Should -Invoke Connect-CMSite -ModuleName AppPackagerCommon -Times 0 -Exactly
        }

        It 'refuses a title with no identifier and no fallback, and takes the fallback when there is one' {
            $path = New-TestManifestFile -ApplicationId ''
            $without = Invoke-SupersedenceApply -ManifestPath $path -SiteCode $script:Site -Settings (New-TestSettings)
            $without.Ok | Should -BeFalse
            $without.Message | Should -Match 'no identity marker'
            Should -Invoke Connect-CMSite -ModuleName AppPackagerCommon -Times 0 -Exactly
            (Invoke-SupersedenceApply -ManifestPath $path -SiteCode $script:Site -Settings (New-TestSettings) -FallbackApplicationId 'catalog:vendor-app').Ok | Should -BeTrue
        }

        It 'refuses a word of the mode that is not allowed, with its place' {
            $settings = New-TestSettings
            $settings.Mode = 'Sometimes'
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings $settings
            $result.Ok | Should -BeFalse
            $result.Message | Should -Match "supersedence mode 'Sometimes' in the site default"
        }
    }

    Context 'the step' {
        It 'sets one relation by name and reports the row as applied' {
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings (New-TestSettings -Mode 'Replace')
            $result.Ok | Should -BeTrue
            $result.Status | Should -Be 'Applied'
            $result.RowStatus | Should -Be 'Supersedence applied'
            $result.Title | Should -Be 'Vendor App 2.0'
            $result.Message | Should -Be 'Vendor App 2.0 supersedes Vendor App 1.0 1.0.0'
            Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter {
                $Name -eq 'Vendor App 2.0' -and $CurrentDeploymentTypeName -eq 'Vendor App 2.0 - x64' -and $SupersededApplicationName -eq 'Vendor App 1.0' -and
                $OldDeploymentTypeName -eq 'Vendor App 1.0 - x64' -and $IsUninstall -eq $false
            }
        }

        It 'passes the uninstall option for Replace and uninstall with the rule Exact' {
            $settings = New-TestSettings -Mode 'ReplaceAndUninstall' -Rule 'Exact'
            $settings.InstalledValues[(Get-SupersedenceRecordKey 'catalog:vendor-app' '2.0.0')] = '2.0.0'
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings $settings
            $result.Ok | Should -BeTrue
            $result.Message | Should -Be 'Vendor App 2.0 supersedes Vendor App 1.0 1.0.0 and uninstalls it'
            Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 1 -Exactly -ParameterFilter { $IsUninstall -eq $true }
        }

        It 'adds nothing on a repeat when the relation exists' {
            Mock Get-CMDeploymentTypeSupersedence -ModuleName AppPackagerCommon { @([pscustomobject]@{ LocalizedDisplayName = 'Vendor App 1.0 - x64' }) }
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings (New-TestSettings)
            $result.Status | Should -Be 'Applied'
            Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 0 -Exactly
        }

        It 'reports an application that is not on the site as not applied and writes nothing' {
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile -Name 'Other App 2.0') -SiteCode $script:Site -Settings (New-TestSettings)
            $result.Ok | Should -BeFalse
            $result.RowStatus | Should -Be 'Supersedence not applied'
            $result.Message | Should -Match "'Other App 2\.0' is not on the site"
            Should -Invoke Set-CMApplicationSupersedence -ModuleName AppPackagerCommon -Times 0 -Exactly
        }

        It 'reports a version that the site does not hold as not applied' {
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile -Version '3.0.0') -SiteCode $script:Site -Settings (New-TestSettings)
            $result.Ok | Should -BeFalse
            $result.Message | Should -Match 'not 3\.0\.0'
        }

        It 'counts no older version as done with its own row text' {
            Mock Get-CMApplication -ModuleName AppPackagerCommon {
                [pscustomobject]@{ CI_ID = 2; LocalizedDisplayName = 'Vendor App 2.0'; SoftwareVersion = '2.0.0'; IsExpired = $false; SDMPackageXML = '<none/>' }
            }
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings (New-TestSettings)
            $result.Ok | Should -BeTrue
            $result.Status | Should -Be 'NoOlderVersion'
            $result.RowStatus | Should -Be 'No older version'
            $result.Message | Should -BeLike 'No older version found*'
        }

        It 'reports a failed write as not applied with the reason' {
            Mock Set-CMApplicationSupersedence -ModuleName AppPackagerCommon { throw 'access denied' }
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings (New-TestSettings)
            $result.Ok | Should -BeFalse
            $result.Message | Should -Be 'Vendor App 2.0 - x64 supersedes Vendor App 1.0 - x64: access denied'
        }

        It 'reports a failed connection as not applied and runs no step' {
            Mock Connect-CMSite -ModuleName AppPackagerCommon { $false }
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings (New-TestSettings)
            $result.Ok | Should -BeFalse
            $result.Message | Should -Match 'The connection to the site MCM failed'
            Should -Invoke Get-CMApplication -ModuleName AppPackagerCommon -Times 0 -Exactly
        }

        It 'puts the plan lines and the warnings of the step into the log lines' {
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings (New-TestSettings)
            @($result.Log | Where-Object { $_ -match '^Supersedence mode\s+: Replace \(site default\)$' }).Count | Should -Be 1
        }

        It 'leaves the location where it found it, also when the step fails' {
            $before = (Get-Location).Path
            Mock Connect-CMSite -ModuleName AppPackagerCommon { Set-Location -LiteralPath $env:TEMP; $true }
            Mock Get-CMApplication -ModuleName AppPackagerCommon { throw 'provider down' }
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings (New-TestSettings)
            $result.Ok | Should -BeFalse
            $result.Message | Should -Match 'lookup of the older version failed: provider down'
            (Get-Location).Path | Should -Be $before
        }
    }

    Context 'the title' {
        It 'resolves the title mode as a publish does, and puts the environment back' {
            $env:APP_PACKAGER_TITLE_MODE = 'IncludeVersion'
            try {
                $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile -Name 'Vendor App 2.0.0') -SiteCode $script:Site -Settings (New-TestSettings) -TitleMode 'NoVersion'
                $result.Title | Should -Be 'Vendor App'
                $env:APP_PACKAGER_TITLE_MODE | Should -Be 'IncludeVersion'
            }
            finally { Remove-Item -LiteralPath 'Env:\APP_PACKAGER_TITLE_MODE' -ErrorAction SilentlyContinue }
        }

        It 'adds the version to the title with the mode IncludeVersion' {
            $result = Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile -Name 'Vendor App' -Version '2.0.0') -SiteCode $script:Site -Settings (New-TestSettings) -TitleMode 'IncludeVersion'
            $result.Title | Should -Be 'Vendor App - 2.0.0'
        }

        It 'removes no variable that was not set before' {
            Remove-Item -LiteralPath 'Env:\APP_PACKAGER_TITLE_MODE' -ErrorAction SilentlyContinue
            [void](Invoke-SupersedenceApply -ManifestPath (New-TestManifestFile) -SiteCode $script:Site -Settings (New-TestSettings) -TitleMode 'NoVersion')
            Test-Path -LiteralPath 'Env:\APP_PACKAGER_TITLE_MODE' | Should -BeFalse
        }
    }
}

Describe 'The sidebar availability of Apply Supersedence' {
    BeforeAll {
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
        $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-SidebarSystemState' }, $true) | Select-Object -First 1
        . ([scriptblock]::Create($fn.Extent.Text))
        function New-SidebarPrefs {
            param([bool]$ConfigMgr = $true, [bool]$Console = $true, [string]$Site = 'MCM', [string]$Mode = 'Replace', $Overrides = $null)
            [pscustomobject]@{
                Systems       = [pscustomobject]@{ ConfigMgr = $ConfigMgr; Intune = $false; Wsus = $false }
                SiteCode      = $Site
                DetectedTools = [pscustomobject]@{ ConfigMgrConsole = [pscustomobject]@{ Found = $Console } }
                Intune        = [pscustomobject]@{ TenantId = ''; ClientId = ''; ClientSecretProtected = ''; DeploymentTarget = 'MECM' }
                Wsus          = [pscustomobject]@{ ServerName = '' }
                Supersedence  = [pscustomobject]@{ Mode = $Mode; TitleOverrides = $(if ($Overrides) { $Overrides } else { [pscustomobject]@{} }) }
            }
        }
    }

    It 'is available with ConfigMgr in use, a site code and a mode other than Off' {
        (Get-SidebarSystemState -Prefs (New-SidebarPrefs)).ApplySupersedenceEnabled | Should -BeTrue
        (Get-SidebarSystemState -Prefs (New-SidebarPrefs -Mode 'Replace and uninstall')).ApplySupersedenceEnabled | Should -BeTrue
    }

    It 'is disabled and says why when ConfigMgr is not in use' {
        $state = Get-SidebarSystemState -Prefs (New-SidebarPrefs -ConfigMgr $false)
        $state.ApplySupersedenceEnabled | Should -BeFalse
        $state.ApplySupersedenceToolTip | Should -Match 'ConfigMgr is not in use'
        $state.ApplySupersedenceToolTip | Should -Match 'Systems in Use'
    }

    It 'is disabled without the console or the site code, with the reason of the ConfigMgr publish' {
        (Get-SidebarSystemState -Prefs (New-SidebarPrefs -Console $false)).ApplySupersedenceToolTip | Should -Match 'console'
        $state = Get-SidebarSystemState -Prefs (New-SidebarPrefs -Site '')
        $state.ApplySupersedenceEnabled | Should -BeFalse
        $state.ApplySupersedenceToolTip | Should -Match 'site code'
    }

    It 'is disabled and says where to turn the mode on when the mode is Off' {
        $state = Get-SidebarSystemState -Prefs (New-SidebarPrefs -Mode 'Off')
        $state.ApplySupersedenceEnabled | Should -BeFalse
        $state.ApplySupersedenceToolTip | Should -Be 'Apply Supersedence: Supersedence is Off. Set the supersedence mode in Options, Packaging, Supersedence.'
    }

    It 'stays available with the mode Off when the override of one title turns it on' {
        $overrides = [pscustomobject]@{ 'catalog:vendor-app' = [pscustomobject]@{ Mode = 'Replace' }; 'catalog:other' = [pscustomobject]@{ VersionRule = 'Exact' } }
        (Get-SidebarSystemState -Prefs (New-SidebarPrefs -Mode 'Off' -Overrides $overrides)).ApplySupersedenceEnabled | Should -BeTrue
        $off = [pscustomobject]@{ 'catalog:vendor-app' = [pscustomobject]@{ Mode = 'Off' } }
        (Get-SidebarSystemState -Prefs (New-SidebarPrefs -Mode 'Off' -Overrides $off)).ApplySupersedenceEnabled | Should -BeFalse
    }

    It 'treats preferences without a Supersedence section as Off' {
        $prefs = New-SidebarPrefs
        $prefs.PSObject.Properties.Remove('Supersedence')
        (Get-SidebarSystemState -Prefs $prefs).ApplySupersedenceEnabled | Should -BeFalse
    }
}

Describe 'The Apply Supersedence button' {
    BeforeAll {
        $script:Shell = [IO.File]::ReadAllText((Join-Path $script:RepoRoot 'start-apppackager.ps1'))
        $script:Xaml = [IO.File]::ReadAllText((Join-Path $script:RepoRoot 'MainWindow.xaml'))
    }

    It 'is a sidebar button in Title Case with a tooltip, after Publish to ConfigMgr' {
        $script:Xaml | Should -Match '<Button x:Name="btnApplySupersedence" Content="Apply Supersedence" Style="\{StaticResource WorkflowButtonStyle\}"\s+ToolTip="[^"]+"'
        $script:Xaml.IndexOf('btnPackage') | Should -BeLessThan $script:Xaml.IndexOf('btnApplySupersedence')
        $script:Xaml.IndexOf('btnApplySupersedence') | Should -BeLessThan $script:Xaml.IndexOf('btnPublishIntune')
    }

    It 'is registered for the theme, for the sidebar availability and for the busy state' {
        $script:Shell | Should -Match '\$script:WorkflowButtons = @\([^)]*\$btnApplySupersedence'
        $script:Shell | Should -Match "@\(\`$btnApplySupersedence, 'ApplySupersedence'\)"
        $script:Shell | Should -Match '\$btnApplySupersedence\.IsEnabled = \$Enabled'
    }

    It 'runs the engine entry point on the checked rows behind the busy display and restores the buttons' {
        $script:Shell | Should -Match '\$btnApplySupersedence\.Add_Click\(\{ Start-ApplySupersedence \}\)'
        $body = [regex]::Match($script:Shell, '(?s)function Start-ApplySupersedence \{.*?\r?\n\}\r?\n').Value
        $body | Should -Match 'Invoke-SupersedenceApply'
        $body | Should -Match 'Get-SelectedRows'
        $body | Should -Match 'Test-PipelineBusy'
        $body | Should -Match 'Show-BusyIndicator'
        $body.IndexOf('Hide-BusyIndicator') | Should -BeGreaterThan $body.IndexOf('finally')
        $body.IndexOf('Set-ActionButtonsEnabled -Enabled $true') | Should -BeGreaterThan $body.IndexOf('finally')
    }

    It 'adds no second copy of the lookup or the step to the window code' {
        $body = [regex]::Match($script:Shell, '(?s)function Start-ApplySupersedence \{.*?\r?\n\}\r?\n').Value
        $body | Should -Not -Match 'Invoke-SupersedenceStep\b|Set-CMApplicationSupersedence|Get-SupersedenceLineApplication'
    }
}
