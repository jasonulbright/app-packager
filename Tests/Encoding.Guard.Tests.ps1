#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Every tracked text file is plain ASCII: no UTF-8 BOM and no byte above 127.

.DESCRIPTION
    A BOM, a dash, a curly quote or an emoji in a script breaks Windows PowerShell 5.1, which reads a file without a BOM
    as ANSI. The rule covers the PowerShell sources and every file that PowerShell or a test reads: JSON, XML, XAML, CSV,
    text, packager definitions, templates and the Markdown that the documentation tests parse. Only binary files are
    exempt. A character outside ASCII is written with a code, for example [char]0x00B7, or with a JSON escape in JSON.
#>

BeforeDiscovery {
    $root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $binary = '.png', '.jpg', '.jpeg', '.gif', '.ico', '.zip', '.dll', '.exe', '.pdf', '.woff', '.woff2', '.ttf', '.msi', '.cab', '.7z', '.nupkg', '.snk', '.pfx', '.cer', '.pdb'
    $tracked = $null
    if (Get-Command git -ErrorAction SilentlyContinue) { $tracked = @(git -C $root ls-files 2>$null) }
    if (-not $tracked -or $tracked.Count -eq 0) {
        $tracked = @(Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object { $_.FullName -notmatch ('(^|[' + [regex]::Escape([string][System.IO.Path]::DirectorySeparatorChar) + '/])(\.git|\.claude)([' + [regex]::Escape([string][System.IO.Path]::DirectorySeparatorChar) + '/]|$)') } | ForEach-Object { $_.FullName.Substring($root.Length + 1) })
    }
    $script:FileD = @($tracked | Where-Object { $binary -notcontains [System.IO.Path]::GetExtension([string]$_).ToLowerInvariant() } | ForEach-Object { @{ Path = [string]$_ } })
}

BeforeAll {
    $script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
}

Describe 'Tracked text files are plain ASCII' {
    It 'finds the text files' -ForEach @(@{ Count = $script:FileD.Count }) {
        $Count | Should -BeGreaterThan 100
    }

    It 'has no BOM and no byte above 127 in <Path>' -ForEach $script:FileD {
        $bytes = [System.IO.File]::ReadAllBytes((Join-Path $script:Root $Path))
        $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
        $hasBom | Should -BeFalse -Because 'a UTF-8 BOM is not allowed'
        $bad = New-Object System.Collections.Generic.List[string]
        $line = 1
        for ($i = 0; $i -lt $bytes.Length; $i++) {
            if ($bytes[$i] -eq 10) { $line++ }
            elseif ($bytes[$i] -gt 127 -and $bad.Count -lt 5) { $bad.Add(('line {0}, byte 0x{1:X2}' -f $line, $bytes[$i])) }
        }
        $bad | Should -BeNullOrEmpty -Because 'write a character outside ASCII with a code such as [char]0x00B7'
    }
}
