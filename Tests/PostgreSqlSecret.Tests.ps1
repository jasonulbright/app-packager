BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking

    function Invoke-SecretBlock {
        param([string]$Block, [string]$ProgramData, [string]$HostExe)
        $script = Join-Path $ProgramData ('probe-{0}.ps1' -f [guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($script, ($Block + "`r`nWrite-Output ('PW=' + `$pgSuperPassword)`r`n"))
        $saved = $env:ProgramData
        try {
            $env:ProgramData = $ProgramData
            $out = & $HostExe -NoProfile -ExecutionPolicy Bypass -File $script 2>&1
            $exit = $LASTEXITCODE
            $acl = ''
            foreach ($secretFolder in @(Get-ChildItem -LiteralPath $ProgramData -Directory -Filter '*-Secret' -ErrorAction SilentlyContinue)) {
                $acl = (& icacls.exe $secretFolder.FullName) -join "`n"
            }
            [pscustomobject]@{ ExitCode = $exit; Output = @($out | ForEach-Object { [string]$_ }); Acl = $acl }
        }
        finally { $env:ProgramData = $saved }
    }
}

Describe 'PostgreSQL superuser password' {
    It 'is not embedded in any PostgreSQL packager' {
        foreach ($n in 13..17) {
            $text = [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot "Packagers\package-postgresql$n.ps1"))
            $text | Should -Not -Match 'P0stgres'
            $text | Should -Match 'New-PostgreSqlSuperuserSecretBlock'
            $text | Should -Match '\$pgSuperPassword'
        }
    }

    It 'rejects a major version that is not a number' {
        { New-PostgreSqlSuperuserSecretBlock -MajorVersion "17'; calc" } | Should -Throw
    }

    It 'creates a 32 character secret once and reuses it on the next run' {
        $dir = Join-Path $TestDrive 'pd1'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        $block = New-PostgreSqlSuperuserSecretBlock -MajorVersion '17'
        foreach ($exe in @('powershell.exe', 'pwsh.exe')) {
            $cmd = Get-Command $exe -ErrorAction SilentlyContinue
            if (-not $cmd) { continue }
            $sub = Join-Path $dir $exe.Replace('.exe', '')
            New-Item -ItemType Directory -Path $sub -Force | Out-Null
            $first = Invoke-SecretBlock -Block $block -ProgramData $sub -HostExe $cmd.Source
            $first.ExitCode | Should -Be 0
            $pw1 = ($first.Output | Where-Object { $_ -like 'PW=*' }) -replace '^PW=', ''
            $pw1.Length | Should -Be 32
            $first.Acl | Should -Match 'NT AUTHORITY\\SYSTEM'
            $first.Acl | Should -Match 'BUILTIN\\Administrators'
            $first.Acl | Should -Not -Match 'BUILTIN\\Users'
            Test-Path -LiteralPath (Join-Path $sub 'PostgreSQL-Superuser-Secret\postgresql-17.dpapi') | Should -BeTrue
            $second = Invoke-SecretBlock -Block $block -ProgramData $sub -HostExe $cmd.Source
            (($second.Output | Where-Object { $_ -like 'PW=*' }) -replace '^PW=', '') | Should -Be $pw1
            $stored = [System.IO.File]::ReadAllBytes((Join-Path $sub 'PostgreSQL-Superuser-Secret\postgresql-17.dpapi'))
            [System.Text.Encoding]::UTF8.GetString($stored) | Should -Not -Match ([regex]::Escape($pw1))
        }
    }

    It 'refuses a folder that a standard user could have created' {
        $dir = Join-Path $TestDrive 'planted'
        $planted = Join-Path $dir 'PostgreSQL-Superuser-Secret'
        New-Item -ItemType Directory -Path $planted -Force | Out-Null
        # An inherited ACL from the test drive grants access to more than SYSTEM, Administrators and this account.
        & icacls.exe $planted /grant 'BUILTIN\Users:(OI)(CI)R' | Out-Null
        [System.IO.File]::WriteAllBytes((Join-Path $planted 'postgresql-17.dpapi'), [byte[]](1, 2, 3))
        $r = Invoke-SecretBlock -Block (New-PostgreSqlSuperuserSecretBlock -MajorVersion '17') -ProgramData $dir -HostExe (Get-Process -Id $PID).Path
        $r.ExitCode | Should -Be 1
        (($r.Output -join ' ') -replace '\s+', ' ') | Should -Match 'Could not create the PostgreSQL superuser secret'
        @($r.Output | Where-Object { $_ -like 'PW=*' }).Count | Should -Be 0
    }
    It 'gives two targets different secrets' {
        $block = New-PostgreSqlSuperuserSecretBlock -MajorVersion '16'
        $hostExe = (Get-Process -Id $PID).Path
        $a = Join-Path $TestDrive 'ta'; $b = Join-Path $TestDrive 'tb'
        New-Item -ItemType Directory -Path $a, $b -Force | Out-Null
        $ra = Invoke-SecretBlock -Block $block -ProgramData $a -HostExe $hostExe
        $rb = Invoke-SecretBlock -Block $block -ProgramData $b -HostExe $hostExe
        ($ra.Output | Where-Object { $_ -like 'PW=*' }) | Should -Not -Be ($rb.Output | Where-Object { $_ -like 'PW=*' })
    }

    It 'exits 1 without a password when the secret cannot be stored' {
        $hostExe = (Get-Process -Id $PID).Path
        $blocker = Join-Path $TestDrive 'blocker'
        New-Item -ItemType Directory -Path $blocker -Force | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $blocker 'PostgreSQL-Superuser-Secret'), 'file where the folder must go')
        $r = Invoke-SecretBlock -Block (New-PostgreSqlSuperuserSecretBlock -MajorVersion '15') -ProgramData $blocker -HostExe $hostExe
        $r.ExitCode | Should -Be 1
        @($r.Output | Where-Object { $_ -like 'PW=*' }).Count | Should -Be 0
    }
}

Describe 'MariaDB root password' {
    It 'generates a root password on the target when none is supplied' {
        $text = [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot 'Packagers\package-mariadb-server.ps1'))
        $text | Should -Match ([regex]::Escape("PASSWORD=__GENERATED__"))
        $text | Should -Match 'New-TargetSecretBlock -FolderName ''MariaDB-Root-Secret'''
        $text | Should -Not -Match 'blank root password'
    }

    It 'passes a 32 character generated password to msiexec' {
        $wrapper = New-MsiWrapperContent -MsiFileName 'mariadb.msi' -ExtraInstallArgs @('SERVICENAME=MariaDB', 'PASSWORD=__GENERATED__')
        $install = $wrapper.Install.Replace('$args += ''PASSWORD=__GENERATED__''', '$args += (''PASSWORD='' + $mariaRootPassword)')
        $install | Should -Not -Match '__GENERATED__'
        $script = (New-TargetSecretBlock -FolderName 'MariaDB-Root-Secret' -FileName 'mariadb.dpapi' -VariableName 'mariaRootPassword' -Label 'MariaDB root') + "`r`n" + $install
        $dir = Join-Path $TestDrive 'maria'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        $record = Join-Path $dir 'args.txt'
        $stub = 'function Start-Process { param($FilePath, $ArgumentList, [switch]$Wait, [switch]$PassThru, [switch]$NoNewWindow) [System.IO.File]::WriteAllText($env:PROBE_ARGS, ($ArgumentList -join "`n")); [pscustomobject]@{ ExitCode = 0 } }'
        $path = Join-Path $dir 'install.ps1'
        [System.IO.File]::WriteAllText($path, $stub + "`r`n" + $script)
        $saved = $env:ProgramData
        try {
            $env:ProgramData = $dir
            $env:PROBE_ARGS = $record
            $r = Invoke-SecretBlock -Block ([System.IO.File]::ReadAllText($path)) -ProgramData $dir -HostExe (Get-Process -Id $PID).Path
        }
        finally { $env:ProgramData = $saved; Remove-Item Env:\PROBE_ARGS -ErrorAction SilentlyContinue }
        $r.ExitCode | Should -Be 0
        $recorded = [System.IO.File]::ReadAllText($record) -split "`n"
        ($recorded | Where-Object { $_ -match '^PASSWORD=[A-Za-z0-9]{32}$' }).Count | Should -Be 1
        $recorded | Should -Contain 'SERVICENAME=MariaDB'
    }
}
