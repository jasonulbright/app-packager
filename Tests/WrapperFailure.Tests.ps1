#Requires -Modules Pester

<#
.SYNOPSIS
    Generated wrappers fail when the operation they run fails.

.DESCRIPTION
    The AppX and file cmdlets report most failures as non-terminating errors,
    and a wrapper that ends in "exit 0" then reports success for work that did
    not happen. Each generated script is run as a child process with the
    cmdlet replaced by a stub that raises a non-terminating error.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking

    function Get-PackagerAst {
        param([string]$Name)
        [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot "Packagers\package-$Name.ps1"), [ref]$null, [ref]$null)
    }
    function Get-FunctionText {
        param([string]$Packager, [string]$Function)
        $node = (Get-PackagerAst $Packager).Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Function }, $true)
        if (-not $node) { throw "package-$Packager.ps1 no longer defines $Function." }
        $node.Extent.Text
    }
    function Get-GeneratedText {
        # Evaluates one assignment of a packager in a scope that supplies its inputs.
        param([string]$Packager, [string]$Variable, [hashtable]$Inputs)
        $ast = Get-PackagerAst $Packager
        $node = $ast.Find({ param($n)
            $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.Left.VariablePath.UserPath -eq $Variable
        }, $true)
        if (-not $node) { throw "package-$Packager.ps1 no longer assigns `$$Variable." }
        $scope = { param($Inputs, $Text) foreach ($k in $Inputs.Keys) { Set-Variable -Name $k -Value $Inputs[$k] }; Invoke-Expression $Text; Get-Variable -Name $Variable -ValueOnly }
        & $scope $Inputs $node.Extent.Text
    }
    function Invoke-Wrapper {
        param([string]$Script, [string]$Stubs, [string]$Folder)
        New-Item -ItemType Directory -Path $Folder -Force | Out-Null
        $path = Join-Path $Folder 'wrapper.ps1'
        [System.IO.File]::WriteAllText($path, ($Stubs + "`r`n" + $Script), (New-Object System.Text.UTF8Encoding($false)))
        $out = & (Get-Command powershell.exe).Source -NoProfile -ExecutionPolicy Bypass -File $path 2>&1
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = (@($out) -join "`n") }
    }
    $script:StubParameters = '[CmdletBinding()] param([switch]$Online, [switch]$SkipLicense, [switch]$AllUsers, $PackagePath, $Regions, $PackageName, $Package, $Path, $Name, [Parameter(ValueFromPipeline)]$InputObject)'
    $script:Fail = 'function __NAME__ { ' + $script:StubParameters + ' process { Write-Error ''injected failure'' } }'
    $script:Pass = 'function __NAME__ { ' + $script:StubParameters + ' process { } }'
    $script:AppxNames = 'Add-AppxProvisionedPackage', 'Remove-AppxProvisionedPackage', 'Remove-AppxPackage', 'Add-AppxPackage'
    $script:AppxStubs = {
        param([string]$Mode)
        $template = if ($Mode -eq 'Fail') { $script:Fail } else { $script:Pass }
        $lines = @($script:AppxNames | ForEach-Object { $template.Replace('__NAME__', $_) })
        # Package queries return one package so the removal calls run.
        $lines += 'function Get-AppxProvisionedPackage { [CmdletBinding()] param([switch]$Online, [switch]$AllUsers, $Name) [pscustomobject]@{ DisplayName = ''Fixture.App''; PackageName = ''Fixture.App_1.0_x64__abc'' } }'
        $lines += 'function Get-AppxPackage { [CmdletBinding()] param([switch]$Online, [switch]$AllUsers, $Name) [pscustomobject]@{ Name = ''Fixture.App''; PackageFullName = ''Fixture.App_1.0_x64__abc'' } }'
        $lines -join "`r`n"
    }
}

Describe 'AppX wrappers' {
    It 'fails the Slack install and removal when the cmdlet fails, and succeeds otherwise' {
        $inputs = @{ MsixFileName = 'slack.msix'; thumbprint = 'AB'; identityName = 'Fixture.App' }
        $install = Get-GeneratedText 'slack' 'installScript' $inputs
        $uninstall = Get-GeneratedText 'slack' 'uninstallScript' $inputs
        $folder = Join-Path $TestDrive 'slack'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $folder 'slack.msix') -Value 'x'
        $sigStub = 'function Get-AuthenticodeSignature { [pscustomobject]@{ Status = ''Valid''; SignerCertificate = [pscustomobject]@{ Thumbprint = ''AB'' } } }'

        (Invoke-Wrapper -Script $install -Stubs ((& $script:AppxStubs 'Fail') + "`r`n" + $sigStub) -Folder $folder).ExitCode | Should -Be 1
        (Invoke-Wrapper -Script $install -Stubs ((& $script:AppxStubs 'Pass') + "`r`n" + $sigStub) -Folder $folder).ExitCode | Should -Be 0
        (Invoke-Wrapper -Script $uninstall -Stubs (& $script:AppxStubs 'Fail') -Folder $folder).ExitCode | Should -Be 1
        (Invoke-Wrapper -Script $uninstall -Stubs (& $script:AppxStubs 'Pass') -Folder $folder).ExitCode | Should -Be 0
    }

    It 'fails the Oh My Posh removal when the cmdlet fails' {
        $inputs = @{ identityName = 'Fixture.App' }
        $uninstall = Get-GeneratedText 'ohmyposh' 'uninstallScript' $inputs
        $folder = Join-Path $TestDrive 'omp'
        (Invoke-Wrapper -Script $uninstall -Stubs (& $script:AppxStubs 'Fail') -Folder $folder).ExitCode | Should -Be 1
        (Invoke-Wrapper -Script $uninstall -Stubs (& $script:AppxStubs 'Pass') -Folder $folder).ExitCode | Should -Be 0
    }

    It 'fails the shared provisioned MSIX install when the cmdlet fails' {
        $folder = Join-Path $TestDrive 'msix'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $folder 'a.msix') -Value 'x'
        $install = (New-MsixWrapperContent -MsixFileName 'a.msix' -Provisioned $true).Install
        (Invoke-Wrapper -Script $install -Stubs (& $script:AppxStubs 'Fail') -Folder $folder).ExitCode | Should -Be 1
        (Invoke-Wrapper -Script $install -Stubs (& $script:AppxStubs 'Pass') -Folder $folder).ExitCode | Should -Be 0
    }

    It 'fails the Teams removal fallback when the cmdlet fails and exits 0 when nothing is installed' {
        $text = Get-FunctionText 'teams-new' 'New-TeamsUninstallWrapper'
        . ([scriptblock]::Create($text))
        $uninstall = New-TeamsUninstallWrapper -Bootstrapper 'missing-bootstrapper.exe'
        $folder = Join-Path $TestDrive 'teams'
        $stubs = @(
            'function Remove-AppxPackage { [CmdletBinding()] param([Parameter(ValueFromPipeline)]$InputObject, [switch]$AllUsers) process { Write-Error ''injected failure'' } }',
            'function Get-AppxPackage { [CmdletBinding()] param([switch]$Online, [switch]$AllUsers, $Name) [pscustomobject]@{ Name = ''MSTeams'' } }'
        ) -join "`r`n"
        (Invoke-Wrapper -Script $uninstall -Stubs $stubs -Folder $folder).ExitCode | Should -Be 1
        $none = 'function Get-AppxPackage { [CmdletBinding()] param([switch]$Online, [switch]$AllUsers, $Name) }'
        (Invoke-Wrapper -Script $uninstall -Stubs $none -Folder $folder).ExitCode | Should -Be 0
    }
}

Describe 'Sysinternals wrappers' {
    It 'fails the extraction and the removal when the cmdlet fails' {
        $inputs = @{ ZipFileName = 'SysinternalsSuite.zip' }
        $install = Get-GeneratedText 'sysinternals' 'installPs1' $inputs
        $uninstall = Get-GeneratedText 'sysinternals' 'uninstallPs1' $inputs
        $folder = Join-Path $TestDrive 'sysinternals'
        $env:ProgramFiles | Should -Not -BeNullOrEmpty
        $failExpand = 'function Expand-Archive { [CmdletBinding()] param($Path, $DestinationPath, [switch]$Force) Write-Error ''injected failure'' }' + "`r`n" +
            'function New-Item { [CmdletBinding()] param($Path, $ItemType, [switch]$Force) }'
        (Invoke-Wrapper -Script $install -Stubs $failExpand -Folder $folder).ExitCode | Should -Be 1
        $failRemove = 'function Test-Path { $true }' + "`r`n" + 'function Remove-Item { [CmdletBinding()] param($Path, [switch]$Recurse, [switch]$Force) Write-Error ''injected failure'' }'
        (Invoke-Wrapper -Script $uninstall -Stubs $failRemove -Folder $folder).ExitCode | Should -Be 1
        $passRemove = 'function Test-Path { $true }' + "`r`n" + 'function Remove-Item { [CmdletBinding()] param($Path, [switch]$Recurse, [switch]$Force) }'
        (Invoke-Wrapper -Script $uninstall -Stubs $passRemove -Folder $folder).ExitCode | Should -Be 0
    }
}
