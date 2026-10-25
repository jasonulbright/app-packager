#Requires -Modules Pester

<#
.SYNOPSIS
    Child-process arguments that carry a path.

.DESCRIPTION
    Start-Process joins an ArgumentList with spaces and adds no quotes, so a
    path with a space reaches the child as two arguments. These cases read the
    packagers and the launcher and require each path argument to be quoted,
    and run a real child to show what the child receives.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Read = { param($relative) [System.IO.File]::ReadAllText((Join-Path $script:RepoRoot $relative)) }
}

Describe 'Packager path arguments' {
    It 'quotes the Office configuration and download paths' {
        foreach ($name in 'm365apps-x64', 'm365apps-x86', 'm365project-x64', 'm365project-x86', 'm365visio-x64', 'm365visio-x86') {
            $text = & $script:Read "Packagers\package-$name.ps1"
            $text | Should -Match ([regex]::Escape('(''"{0}"'' -f $downloadXmlPath)'))
            $text | Should -Match ([regex]::Escape('"`"$configPath`""'))
            $text | Should -Not -Match ([regex]::Escape('@(''/download'', $downloadXmlPath)'))
            $text | Should -Not -Match ([regex]::Escape('@(''''/configure'''', $configPath)'))
        }
    }

    It 'quotes the layout path for the ADK and the WinPE add-on' {
        foreach ($name in 'windowsadk', 'windowspeaddon') {
            (& $script:Read "Packagers\package-$name.ps1") | Should -Match ([regex]::Escape('''/layout'', (''"{0}"'' -f $LayoutPath)'))
        }
    }

    It 'quotes the Adobe extraction paths and the GeoGebra MSI path' {
        $adobe = & $script:Read 'Packagers\package-adobereader.ps1'
        $adobe | Should -Match ([regex]::Escape('(''"-o{0}"'' -f $localContentPath)'))
        $adobe | Should -Match ([regex]::Escape('(''"{0}"'' -f $localExe)'))
        (& $script:Read 'Packagers\package-geogebra.ps1') | Should -Match ([regex]::Escape('"`"$msiPath`""'))
    }

    It 'quotes the Visual C++ log paths' {
        $text = & $script:Read 'Packagers\package-msvcruntimes.ps1'
        foreach ($log in 'x86.install.log', 'x64.install.log', 'x86.uninstall.log', 'x64.uninstall.log') {
            $text | Should -Match ([regex]::Escape("''`"'' + (Join-Path `$PSScriptRoot ''$log'') + ''`"''"))
        }
    }

    It 'parses every changed packager' {
        $names = 'adobereader', 'geogebra', 'm365apps-x64', 'm365apps-x86', 'm365project-x64', 'm365project-x86', 'm365visio-x64', 'm365visio-x86', 'msvcruntimes', 'windowsadk', 'windowspeaddon'
        foreach ($name in $names) {
            $errors = $null
            [void][System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot "Packagers\package-$name.ps1"), [ref]$null, [ref]$errors)
            @($errors).Count | Should -Be 0 -Because $name
        }
    }
}

Describe 'Launcher path arguments' {
    It 'quotes the relaunch, the update bootstrap and the folders opened in Explorer' {
        $text = & $script:Read 'start-apppackager.ps1'
        $text | Should -Match ([regex]::Escape('''-File'', (''"{0}"'' -f $launcher)'))
        $text | Should -Match ([regex]::Escape('''-File'', (''"{0}"'' -f $bootstrap)'))
        $text | Should -Match ([regex]::Escape('''-File'',''"$($relaunch.Replace("''","''''"))"'''))
        $text | Should -Not -Match 'Start-Process\s+"explorer\.exe"\s+-ArgumentList\s+\$'
    }
}

Describe 'Start-Process with a quoted path' {
    It 'delivers a path with a space to the child as one argument' {
        $folder = Join-Path $TestDrive 'folder with space'
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
        $probe = Join-Path $folder 'probe.ps1'
        Set-Content -LiteralPath $probe -Value 'param([string]$First, [string]$Second) [System.IO.File]::WriteAllText($env:PROBE_OUT, ($First + "|" + $Second))' -Encoding ASCII
        $out = Join-Path $TestDrive 'probe.out'
        $env:PROBE_OUT = $out
        try {
            $proc = Start-Process -FilePath (Get-Command powershell.exe).Source -Wait -PassThru -NoNewWindow `
                -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $probe), '-First', ('"{0}"' -f (Join-Path $folder 'a b.txt')), '-Second', 'plain')
            $proc.ExitCode | Should -Be 0
            [System.IO.File]::ReadAllText($out) | Should -Be ((Join-Path $folder 'a b.txt') + '|plain')
        }
        finally { Remove-Item Env:\PROBE_OUT -ErrorAction SilentlyContinue }
    }
}
