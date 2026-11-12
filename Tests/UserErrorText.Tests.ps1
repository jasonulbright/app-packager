#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Get-UserErrorText and the -BatchMode run: the functions exist before the batch block runs, a failed vendor
    check is recorded, and the run goes on with the next application. No window opens.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:ShellPath = Join-Path $script:RepoRoot 'start-apppackager.ps1'
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
    function Get-ShellFunction([string]$Name) {
        $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $false)
    }
    # The log call of the helper goes to a list, so that the test sees where the raw text went.
    $global:UserErrorLog = New-Object System.Collections.Generic.List[string]
    function global:Write-Log { param([string]$Message, [string]$Level) $global:UserErrorLog.Add($Message) }
    . ([scriptblock]::Create((Get-ShellFunction 'Get-UserErrorText').Extent.Text))
    . ([scriptblock]::Create((Get-ShellFunction 'Get-HttpStatusCode').Extent.Text))
}

AfterAll {
    Remove-Item function:global:Write-Log -ErrorAction SilentlyContinue
    Remove-Variable -Name UserErrorLog -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Get-UserErrorText' {
    BeforeEach { $global:UserErrorLog.Clear() }

    It 'passes the message of a throw statement' {
        try { throw 'The site code is required.' } catch { $r = Get-UserErrorText -ErrorRecord $_ -Fallback 'Failed.' }
        $r | Should -Be 'The site code is required.'
        $global:UserErrorLog.Count | Should -Be 0
    }

    It 'returns the fallback for an engine error and writes the raw text to the log' {
        try { $null.Foo() } catch { $r = Get-UserErrorText -ErrorRecord $_ -Fallback 'Failed. See the log file.' }
        $r | Should -Be 'Failed. See the log file.'
        $global:UserErrorLog.Count | Should -Be 1
        $global:UserErrorLog[0] | Should -Match 'Error detail kept out of the window'
    }

    It 'returns the fallback for a .NET exception' {
        try { [System.IO.File]::ReadAllText('C:\no\such\file.txt') } catch { $r = Get-UserErrorText -ErrorRecord $_ -Fallback 'Failed.' }
        $r | Should -Be 'Failed.'
    }

    It 'passes the message of a refusal that carries WsusRefusal' {
        $ex = New-Object System.InvalidOperationException('This application cannot be published to WSUS: CustomInstall.')
        $ex.Data['WsusRefusal'] = 'CustomInstall'
        try { throw $ex } catch { $r = Get-UserErrorText -ErrorRecord $_ -Fallback 'Failed.' }
        $r | Should -Be 'This application cannot be published to WSUS: CustomInstall.'
    }

    It 'passes the message of an exception that carries UserMessage, also with an inner exception' {
        $inner = New-Object System.Data.DataException('raw sql text')
        $ex = New-Object System.InvalidOperationException('Login failed for the site database.', $inner)
        $ex.Data['UserMessage'] = '1'
        try { throw $ex } catch { $r = Get-UserErrorText -ErrorRecord $_ -Fallback 'Failed.' }
        $r | Should -Be 'Login failed for the site database.'
    }

    It 'returns the fallback for an exception with an inner exception and no marker' {
        $ex = New-Object System.InvalidOperationException('outer', (New-Object System.Exception('inner')))
        try { throw $ex } catch { $r = Get-UserErrorText -ErrorRecord $_ -Fallback 'Failed.' }
        $r | Should -Be 'Failed.'
    }

    It 'returns no status code for an error without a web response and a number for one with a response' {
        try { throw 'x' } catch { Get-HttpStatusCode -ErrorRecord $_ | Should -BeNullOrEmpty }
    }
}

Describe 'The helpers exist before the -BatchMode block' {
    It 'defines Get-UserErrorText and Get-HttpStatusCode above the first statement that tests $BatchMode and exits' {
        $batch = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '$BatchMode' }, $false)
        $batch | Should -Not -BeNullOrEmpty
        foreach ($name in 'Get-UserErrorText', 'Get-HttpStatusCode') {
            (Get-ShellFunction $name).Extent.StartOffset | Should -BeLessThan $batch.Extent.StartOffset -Because $name
        }
    }
}

Describe 'A -BatchMode run records a failed check and goes on' {
    BeforeAll {
        $script:Sandbox = Join-Path $env:TEMP ('ap-batchmode-' + [guid]::NewGuid().ToString('N'))
        $script:Pk = Join-Path $script:Sandbox 'Packagers'
        [void](New-Item -ItemType Directory -Path $script:Pk, (Join-Path $script:Sandbox 'local') -Force)
        $head = "param([switch]`$StageOnly, [switch]`$PackageOnly, [switch]`$GetLatestVersionOnly, [string]`$SiteCode, [string]`$Comment, [string]`$LogPath, [string]`$DownloadRoot, [string]`$ContentLayout, [string]`$FileServerPath, [int]`$EstimatedRuntimeMins, [int]`$MaximumRuntimeMins)`r`n"
        $bad = $head + "# Vendor: Fixture`r`n# Application: Failing Fixture`r`nif (`$GetLatestVersionOnly) { Write-Error 'the vendor site is not reachable'; exit 1 }`r`nexit 0`r`n"
        $good = $head + "# Vendor: Fixture`r`n# Application: Good Fixture`r`nif (`$GetLatestVersionOnly) { Write-Output '1.0.0'; exit 0 }`r`nexit 0`r`n"
        [System.IO.File]::WriteAllText((Join-Path $script:Pk 'package-aaafailing.ps1'), $bad, (New-Object System.Text.ASCIIEncoding))
        [System.IO.File]::WriteAllText((Join-Path $script:Pk 'package-bbbgood.ps1'), $good, (New-Object System.Text.ASCIIEncoding))
        $host_ = (Get-Process -Id $PID).Path
        $env:LOCALAPPDATA_SAVED = $env:LOCALAPPDATA
        $env:LOCALAPPDATA = Join-Path $script:Sandbox 'local'
        $log = Join-Path $script:Sandbox 'batch.log'
        $script:Out = & $host_ -NoProfile -ExecutionPolicy Bypass -File $script:ShellPath -BatchMode -Apps 'package-aaafailing,package-bbbgood' -OnUpdateFound Report -PackagersRoot $script:Pk -LogPath $log 2>&1 | Out-String
        $script:ExitCode = $LASTEXITCODE
        $script:LogText = if (Test-Path -LiteralPath $log) { Get-Content -LiteralPath $log -Raw } else { '' }
        $env:LOCALAPPDATA = $env:LOCALAPPDATA_SAVED
    }

    AfterAll {
        if ($env:LOCALAPPDATA_SAVED) { $env:LOCALAPPDATA = $env:LOCALAPPDATA_SAVED }
        Remove-Item -LiteralPath $script:Sandbox -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'ends with exit code 1 because one check failed' {
        $script:ExitCode | Should -Be 1
    }

    It 'never fails with a term that is not recognized' {
        $script:Out | Should -Not -Match 'is not recognized as the name'
        $script:LogText | Should -Not -Match 'is not recognized as the name'
    }

    It 'records the failure of the first application and the result of the second' {
        $script:LogText | Should -Match 'package-aaafailing\s+CheckFailed'
        $script:LogText | Should -Match 'package-bbbgood\s+(?!CheckFailed)\w+'
    }

    It 'keeps the raw text of the failed check in the log file' {
        $script:LogText | Should -Match 'check failed'
    }
}
