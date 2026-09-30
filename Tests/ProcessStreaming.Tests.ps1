BeforeAll {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-ProcessWithStreaming' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))

    function New-ChildStartInfo {
        param([Parameter(Mandatory)][string]$Command)
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'powershell.exe'
        $psi.Arguments = ('-NoProfile -NonInteractive -Command "{0}"' -f $Command)
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        return $psi
    }
}

Describe 'Invoke-ProcessWithStreaming' {
    It 'returns one result object that accepts a post-step note' {
        $result = Invoke-ProcessWithStreaming -StartInfo (New-ChildStartInfo 'Start-Sleep -Milliseconds 400; Write-Output done') -OutLog (Join-Path $TestDrive 'out.log') -ErrLog (Join-Path $TestDrive 'err.log')

        @($result).Count | Should -Be 1
        $result.ExitCode | Should -Be 0
        $result.StdOut | Should -Be 'done'
        $result | Add-Member -NotePropertyName WsusPublish -NotePropertyValue ([pscustomobject]@{ Ok = $true }) -Force
        $result.PSObject.Properties['WsusPublish'] | Should -Not -BeNullOrEmpty
    }

    It 'returns one result object when the idle timeout kills the child' {
        $result = Invoke-ProcessWithStreaming -StartInfo (New-ChildStartInfo 'Start-Sleep -Seconds 20') -OutLog (Join-Path $TestDrive 'idle-out.log') -ErrLog (Join-Path $TestDrive 'idle-err.log') -IdleTimeoutSeconds 1

        @($result).Count | Should -Be 1
        $result.StdOut | Should -Match 'idle for 1 seconds; killed'
    }
}
