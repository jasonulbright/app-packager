#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Parse-only tests of the capture tool: the lab lock, the own folder and task, the cleanup order. Nothing runs.

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:CapturePath = Join-Path $script:Root 'tools\guide-capture\Invoke-GuideCapture.ps1'
    $script:Text = Get-Content -LiteralPath $script:CapturePath -Raw
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:CapturePath, [ref]$null, [ref]$null)
    $script:Code = [regex]::Replace($script:Text, '(?s)<#.*?#>', '')
}

Describe 'The capture tool takes the lab lock and keeps to its own folder' {
    It 'parses without an error' {
        $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($script:CapturePath, [ref]$null, [ref]$errors)
        @($errors) | Should -BeNullOrEmpty
    }

    It 'works in its own folder with its own task name and never in C:\guide\ui' {
        $script:Code | Should -Match "\`$app = 'app-packager'"
        $script:Code | Should -Match '\$root = "C:\\guide\\\$app"'
        $script:Code | Should -Match '\$taskName = "GuideShots-\$app"'
        $script:Code | Should -Not -Match 'C:\\guide\\ui'
        $script:Code | Should -Not -Match 'C:\\guide\\dl'
    }

    It 'never deletes the C:\guide folder' {
        $script:Code | Should -Not -Match "Remove-Item\s+['""]?C:\\guide['""]?(\s|$)"
        $script:Code | Should -Not -Match 'Remove-Item[^\r\n]*\$lockPath[^\r\n]*-Recurse'
    }

    It 'takes the lock with New-Item and an IOException catch' {
        $script:Code | Should -Match "New-Item -ItemType File -Path \`$path -ErrorAction Stop"
        $script:Code | Should -Match 'catch \[System\.IO\.IOException\]'
        $script:Code | Should -Match "\`$lockPath = 'C:\\guide\\capture\.lock'"
    }

    It 'renames a stale lock to .stale.<token> and reads the age of the renamed file again' {
        $script:Code | Should -Match "'\{0\}\.stale\.\{1\}' -f \`$path, \`$token"
        $script:Code | Should -Match 'Move-Item -LiteralPath \$path -Destination \$stale'
        $script:Code | Should -Match '\$moved\.LastWriteTime'
        $script:Code | Should -Match '-gt 60|\$staleMinutes'
        $script:Code | Should -Match '60, \$LockWaitMinutes|, 60,'
    }

    It 'sleeps in every retry of the lock loop, so that the loop never spins' {
        $loop = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.WhileStatementAst] }, $true)
        $loop | Should -Not -BeNullOrEmpty
        $loop.Extent.Text | Should -Match 'Start-Sleep -Seconds 15'
        $loop.Extent.Text | Should -Not -Match '\bcontinue\b'
    }

    It 'touches the lock while it polls and removes only its own lock' {
        $script:Code | Should -Match '\.LastWriteTime = Get-Date'
        $script:Code | Should -Match '-eq \$token'
    }

    It 'handles an empty lock file as a lock' {
        # The holder creates the file first and writes its token second. An empty file is a lock, never an error.
        $script:Code | Should -Match '\[string\]\(Get-Content -LiteralPath \$path -Raw'
        $script:Code | Should -Match 'Set-Content -LiteralPath \$path -Value \$token'
    }

    It 'opens the session inside the try block' {
        $try = $script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] }, $false) | Select-Object -First 1
        $try.Body.Extent.Text | Should -Match 'New-PSSession'
        $script:Code | Should -Not -Match '(?m)^\$session = New-PSSession'
    }

    It 'cleans up in order: stop the task, stop the processes, unregister, delete the folder, release the lock' {
        $try = $script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] }, $false) | Select-Object -First 1
        $finally = $try.Finally.Extent.Text
        $order = @('Stop-ScheduledTask', 'Stop-Process', 'Unregister-ScheduledTask', 'Remove-Item $r', 'Remove-Item -LiteralPath $path') | ForEach-Object { $finally.IndexOf($_) }
        $order | Should -Not -Contain -1
        $sorted = @($order | Sort-Object)
        ($order -join ',') | Should -Be ($sorted -join ',')
    }

    It 'filters the copied files by names at the repository root and leaves out .claude' {
        $script:Code | Should -Match "-notin '\.git', '\.claude'"
        $script:Code | Should -Match 'Get-ChildItem -LiteralPath \$repo -Force'
    }
}
