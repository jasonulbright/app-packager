#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The confirmation before the supersedence over a Package run of two applications: the answer, Remember, Skip and No.

.DESCRIPTION
    Get-SupersedenceRunAnswer is read from start-apppackager.ps1 and run with stubs for the site and the operator. The
    operator's answer is given by the stub of Start-Sleep, which is what the handshake with the window waits on.
#>

BeforeAll {
    $script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Root 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-SupersedenceRunAnswer' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))

    # A drive that stands for the site drive.
    New-PSDrive -Name TST -PSProvider FileSystem -Root $env:TEMP -Scope Global | Out-Null

    function New-ApplicationId { param($Kind, $ScriptPath) 'catalog:' + [System.IO.Path]::GetFileNameWithoutExtension($ScriptPath) }
    function Get-SupersedenceSettings { }
    function Get-SupersedenceCommandLine { }
    function Resolve-Supersedence { param($Settings, $ApplicationId, $CommandLine) [pscustomobject]@{ Mode = [pscustomobject]@{ Value = $script:Mode } } }
    function Test-SupersedenceVersionInTitle { param($Title, $Version) $Title -like "*$Version*" }
    function Find-NewestStageManifestForPackager { param($PackagerPath, $DownloadRoot) $null }
    function Get-SupersedenceBaseTitle { param($AppName, $Version) 'Base' }
    function Connect-CMSite { param($SiteCode, $ProviderMachineName) $true }
    function Find-SupersedenceOldApplication {
        param($BaseTitle, $Vendor, $NewApplicationName, $NewVersion)
        $script:Lookups++
        [pscustomobject]@{ Old = $(if ($script:HasOld) { [pscustomobject]@{ Title = 'Old'; SoftwareVersion = '1.0.0' } } else { $null }); Summary = 'none' }
    }
    function Get-UserErrorText { param($ErrorRecord, $Fallback) $ErrorRecord.Exception.Message }
    function New-RunState { @{ LogQueue = (New-Object System.Collections.Queue); SupersedenceRequest = $null; SupersedenceResponse = $null; SupersedenceDecisionForAll = ''; CancelRequested = $false; Step = '' } }
    function Invoke-Answer {
        param($State, [string]$Name)
        $args2 = @{ SiteCode = 'TST'; ProviderMachineName = 'cm01'; PackagerPath = "C:\x\package-$Name.ps1"; DownloadRoot = 'C:\x' }
        Get-SupersedenceRunAnswer -State $State -AppLabel $Name -PackageArgs $args2 -Identity ([pscustomobject]@{ AppName = "$Name - 2.0.0"; Version = '2.0.0' })
    }
}

AfterAll {
    Remove-PSDrive -Name TST -Scope Global -ErrorAction SilentlyContinue
}

Describe 'The confirmation over a run of two applications' {
    BeforeEach {
        $script:Mode = 'Replace'
        $script:HasOld = $true
        $script:Lookups = 0
        $script:Asked = 0
        $script:NextAnswer = $null
        $script:State = New-RunState
        # The operator: answers the parked request, as the window does.
        Mock Start-Sleep {
            if ($script:State.SupersedenceRequest) { $script:Asked++; $script:State.SupersedenceResponse = $script:NextAnswer }
        }
    }

    It 'asks for the first application and keeps the answer for the second when Remember is set' {
        $script:NextAnswer = [pscustomobject]@{ Choice = 'Yes'; Remember = $true }
        (Invoke-Answer -State $script:State -Name 'one').Answer | Should -Be 'Yes'
        $script:State.SupersedenceDecisionForAll | Should -Be 'Yes'
        $second = Invoke-Answer -State $script:State -Name 'two'
        $second.Answer | Should -Be 'Yes'
        $second.Cancel | Should -BeFalse
        $script:Asked | Should -Be 1
        $script:Lookups | Should -Be 1
    }

    It 'keeps a Skip for the run the same way' {
        $script:NextAnswer = [pscustomobject]@{ Choice = 'Skip'; Remember = $true }
        (Invoke-Answer -State $script:State -Name 'one').Answer | Should -Be 'Skip'
        (Invoke-Answer -State $script:State -Name 'two').Answer | Should -Be 'Skip'
        $script:Asked | Should -Be 1
    }

    It 'asks again for the second application when Remember is not set' {
        $script:NextAnswer = [pscustomobject]@{ Choice = 'Yes'; Remember = $false }
        (Invoke-Answer -State $script:State -Name 'one').Answer | Should -Be 'Yes'
        $script:State.SupersedenceDecisionForAll | Should -BeNullOrEmpty
        (Invoke-Answer -State $script:State -Name 'two').Answer | Should -Be 'Yes'
        $script:Asked | Should -Be 2
        $script:Lookups | Should -Be 2
    }

    It 'stops the run on No and asks nothing more' {
        $script:NextAnswer = [pscustomobject]@{ Choice = 'No'; Remember = $true }
        $first = Invoke-Answer -State $script:State -Name 'one'
        $first.Cancel | Should -BeTrue
        $script:State.SupersedenceDecisionForAll | Should -BeNullOrEmpty
    }

    It 'asks nothing when supersedence is off, and logs one line when there is no older version' {
        $script:Mode = 'Off'
        (Invoke-Answer -State $script:State -Name 'one').Answer | Should -BeNullOrEmpty
        $script:Asked | Should -Be 0
        $script:Lookups | Should -Be 0
        $script:Mode = 'Replace'
        $script:HasOld = $false
        (Invoke-Answer -State $script:State -Name 'two').Answer | Should -BeNullOrEmpty
        $script:Asked | Should -Be 0
        @($script:State.LogQueue.ToArray() | Where-Object { $_ -like 'Supersedence: two: *' }).Count | Should -Be 1
    }
}
