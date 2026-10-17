#Requires -Modules Pester

BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    function Get-ScriptFunctionText([string]$Name) {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
        if (-not $fn) { throw "Function $Name not found in start-apppackager.ps1" }
        $fn.Extent.Text
    }
    . ([scriptblock]::Create((Get-ScriptFunctionText 'Import-BackgroundWorkerModule')))
    $script:LookupText = Get-ScriptFunctionText 'Get-MecmCurrentVersionByCMName'

    # The runspace has no ConfigMgr module: these stand in for the site drive
    # and for a site where only a decorated title exists.
    $script:SiteStubs = @'
function Get-CMApplication {
    [CmdletBinding()]
    param([string]$Name, [switch]$DisableWildcardHandling)
    $all = @(
        [pscustomobject]@{ LocalizedDisplayName = 'M365 Apps for Enterprise (x64) (MEC)'; Name = 'M365 Apps for Enterprise (x64) (MEC)'; SoftwareVersion = '16.0.17928.20156' }
        [pscustomobject]@{ LocalizedDisplayName = 'M365 Apps for Business'; Name = 'M365 Apps for Business'; SoftwareVersion = '16.0.1' }
    )
    $all | Where-Object { $_.LocalizedDisplayName -like $Name }
}
function Get-PSDrive { [CmdletBinding()] param($Name, $PSProvider) [pscustomobject]@{ Name = $Name; Root = 'cm01.contoso.com' } }
function Set-Location { [CmdletBinding()] param($Path) }
function Get-Location { 'C:\' }
'@
}

Describe 'Background worker runspace' {
    BeforeEach {
        $script:Runspace = [runspacefactory]::CreateRunspace()
        $script:Runspace.ApartmentState = 'STA'
        $script:Runspace.ThreadOptions = 'ReuseThread'
        $script:Runspace.Open()
        Import-BackgroundWorkerModule -Runspace $script:Runspace -Root $script:RepoRoot
        $inject = [powershell]::Create()
        $inject.Runspace = $script:Runspace
        [void]$inject.AddScript($script:LookupText + [Environment]::NewLine + $script:SiteStubs)
        [void]$inject.Invoke()
        $inject.Dispose()
    }
    AfterEach {
        $script:Runspace.Dispose()
    }

    It 'resolves the ConfigMgr title rule commands' {
        $ps = [powershell]::Create()
        $ps.Runspace = $script:Runspace
        [void]$ps.AddScript({ @('Test-MecmApplicationTitle', 'Get-MecmApplicationNamePattern') | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) } })
        $missing = @($ps.Invoke())
        $ps.Dispose()
        $missing | Should -BeNullOrEmpty
    }

    It 'finds a decorated title when the exact CMName is absent' {
        $ps = [powershell]::Create()
        $ps.Runspace = $script:Runspace
        [void]$ps.AddScript({ Get-MecmCurrentVersionByCMName -SiteCode 'MCM' -CMName 'M365 Apps for Enterprise' })
        $result = @($ps.Invoke())
        $errors = @($ps.Streams.Error | ForEach-Object { $_.ToString() })
        $ps.Dispose()
        $errors | Should -BeNullOrEmpty
        $result.Count | Should -Be 1
        $result[0].Found | Should -BeTrue
        $result[0].DisplayName | Should -Be 'M365 Apps for Enterprise (x64) (MEC)'
        $result[0].SoftwareVersion | Should -Be '16.0.17928.20156'
    }
}
