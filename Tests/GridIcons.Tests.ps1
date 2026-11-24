#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The main grid shows the packager icons from the icon pack.

.DESCRIPTION
    Extracts the grid icon functions from start-apppackager.ps1 and runs them
    against a scratch icon folder: the lookup by packager name, the placeholder
    for a packager without an icon, the in-memory read (the file can be
    replaced and deleted while the grid holds the icon), the cache reset after
    a pack update, and the markup of the column.
#>

BeforeAll {
    Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Drawing
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:StartPath = Join-Path $script:Root 'start-apppackager.ps1'
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:StartPath, [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Reset-GridIconCache', 'Get-GridIconIndex', 'Get-GridIconPlaceholder', 'Get-GridIconSource', 'Set-GridRowImage') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $script:IconDir = Join-Path ([IO.Path]::GetTempPath()) ('gridicons-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $script:IconDir -Force | Out-Null
    function Get-IconPackRoot { $script:IconDir }

    function New-TestPng {
        param([string]$Path, [int]$Size, [System.Drawing.Color]$Color)
        $bmp = New-Object System.Drawing.Bitmap ($Size, $Size)
        try {
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            try { $g.Clear($Color) } finally { $g.Dispose() }
            $bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally { $bmp.Dispose() }
    }
    function Get-FirstPixel {
        param($Source)
        $px = New-Object byte[] 4
        $Source.CopyPixels((New-Object System.Windows.Int32Rect (0, 0, 1, 1)), $px, 4, 0)
        return ('{0},{1},{2}' -f $px[2], $px[1], $px[0])
    }
}

AfterAll {
    if ($script:IconDir -and (Test-Path -LiteralPath $script:IconDir)) { Remove-Item -LiteralPath $script:IconDir -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Grid icon lookup' {
    BeforeEach {
        Get-ChildItem -LiteralPath $script:IconDir -File | Remove-Item -Force
        Reset-GridIconCache
    }

    It 'finds the icon by the packager name, in any letter case' {
        New-TestPng -Path (Join-Path $script:IconDir 'PowerShell7Msix.png') -Size 256 -Color ([System.Drawing.Color]::FromArgb(255, 200, 0, 0))
        $s = Get-GridIconSource -ScriptName 'package-powershell7msix.ps1'
        $s | Should -BeOfType ([System.Windows.Media.Imaging.BitmapImage])
        $s.PixelWidth | Should -Be 48
        $s.IsFrozen | Should -BeTrue
    }

    It 'decodes the icon small, so 300 large icons do not fill memory' {
        New-TestPng -Path (Join-Path $script:IconDir 'big.png') -Size 512 -Color ([System.Drawing.Color]::Blue)
        (Get-GridIconSource -ScriptName 'package-big.ps1').PixelWidth | Should -Be 48
    }

    It 'shows one shared placeholder for a packager without an icon' {
        $a = Get-GridIconSource -ScriptName 'package-nothere.ps1'
        $b = Get-GridIconSource -ScriptName 'package-alsonot.ps1'
        $a | Should -BeOfType ([System.Windows.Media.DrawingImage])
        $a.IsFrozen | Should -BeTrue
        [object]::ReferenceEquals($a, $b) | Should -BeTrue
    }

    It 'shows the placeholder for an icon file that cannot be decoded' {
        Set-Content -LiteralPath (Join-Path $script:IconDir 'broken.png') -Value 'not an image'
        (Get-GridIconSource -ScriptName 'package-broken.ps1') | Should -BeOfType ([System.Windows.Media.DrawingImage])
    }

    It 'ignores the manifest and other files in the pack folder' {
        Set-Content -LiteralPath (Join-Path $script:IconDir 'manifest.json') -Value '{}'
        (Get-GridIconSource -ScriptName 'package-manifest.ps1') | Should -BeOfType ([System.Windows.Media.DrawingImage])
    }

    It 'prefers the png when the pack holds a png and an ico of one name' {
        New-TestPng -Path (Join-Path $script:IconDir 'both.png') -Size 64 -Color ([System.Drawing.Color]::Green)
        Set-Content -LiteralPath (Join-Path $script:IconDir 'both.ico') -Value 'not an icon'
        (Get-GridIconSource -ScriptName 'package-both.ps1') | Should -BeOfType ([System.Windows.Media.Imaging.BitmapImage])
    }
}

Describe 'Grid icons and the icon pack update' {
    BeforeEach {
        Get-ChildItem -LiteralPath $script:IconDir -File | Remove-Item -Force
        Reset-GridIconCache
    }

    It 'holds no handle on the file: the file can be replaced and deleted while the grid keeps the icon' {
        $file = Join-Path $script:IconDir 'locked.png'
        New-TestPng -Path $file -Size 128 -Color ([System.Drawing.Color]::FromArgb(255, 10, 20, 30))
        $s = Get-GridIconSource -ScriptName 'package-locked.ps1'
        { New-TestPng -Path $file -Size 128 -Color ([System.Drawing.Color]::FromArgb(255, 200, 100, 50)) } | Should -Not -Throw
        { Remove-Item -LiteralPath $file -Force -ErrorAction Stop } | Should -Not -Throw
        Get-FirstPixel $s | Should -Be '10,20,30'
    }

    It 'serves the cached icon until the cache is reset, then reads the new file' {
        $file = Join-Path $script:IconDir 'swap.png'
        New-TestPng -Path $file -Size 128 -Color ([System.Drawing.Color]::FromArgb(255, 10, 20, 30))
        $first = Get-GridIconSource -ScriptName 'package-swap.ps1'
        New-TestPng -Path $file -Size 128 -Color ([System.Drawing.Color]::FromArgb(255, 200, 100, 50))
        [object]::ReferenceEquals($first, (Get-GridIconSource -ScriptName 'package-swap.ps1')) | Should -BeTrue
        Reset-GridIconCache
        Get-FirstPixel (Get-GridIconSource -ScriptName 'package-swap.ps1') | Should -Be '200,100,50'
    }

    It 'finds an icon that the update added after the placeholder was cached' {
        $placeholder = Get-GridIconSource -ScriptName 'package-late.ps1'
        $placeholder | Should -BeOfType ([System.Windows.Media.DrawingImage])
        New-TestPng -Path (Join-Path $script:IconDir 'late.png') -Size 64 -Color ([System.Drawing.Color]::Red)
        Reset-GridIconCache
        Get-GridIconSource -ScriptName 'package-late.ps1' | Should -BeOfType ([System.Windows.Media.Imaging.BitmapImage])
    }
}

Describe 'Grid icon follows the row identity' {
    BeforeEach {
        Get-ChildItem -LiteralPath $script:IconDir -File | Remove-Item -Force
        Reset-GridIconCache
        New-TestPng -Path (Join-Path $script:IconDir 'office.png') -Size 64 -Color ([System.Drawing.Color]::FromArgb(255, 200, 0, 0))
        New-TestPng -Path (Join-Path $script:IconDir 'odbc.png') -Size 64 -Color ([System.Drawing.Color]::FromArgb(255, 0, 0, 200))
    }

    It 'replaces the icon of a recycled container when it shows another packager' {
        $image = New-Object System.Windows.Controls.Image
        $a = [pscustomobject]@{ Script = 'package-office.ps1' }
        $b = [pscustomobject]@{ Script = 'package-odbc.ps1' }
        Set-GridRowImage -Image $image -Item $a
        Get-FirstPixel $image.Source | Should -Be '200,0,0'
        Set-GridRowImage -Image $image -Item $b
        Get-FirstPixel $image.Source | Should -Be '0,0,200'
        Set-GridRowImage -Image $image -Item $a
        Get-FirstPixel $image.Source | Should -Be '200,0,0'
    }

    It 'gives every row of a sorted list its own icon, whatever the order' {
        $rows = @('office', 'odbc', 'nothere') | ForEach-Object { [pscustomobject]@{ Script = "package-$_.ps1" } }
        $containers = 1..3 | ForEach-Object { New-Object System.Windows.Controls.Image }
        foreach ($order in @(@(0, 1, 2), @(2, 1, 0), @(1, 2, 0), @(0, 2, 1))) {
            for ($i = 0; $i -lt 3; $i++) { Set-GridRowImage -Image $containers[$i] -Item $rows[$order[$i]] }
            for ($i = 0; $i -lt 3; $i++) {
                $expected = Get-GridIconSource -ScriptName $rows[$order[$i]].Script
                [object]::ReferenceEquals($containers[$i].Source, $expected) | Should -BeTrue
            }
        }
    }

    It 'keeps the placeholder identity for a packager without an icon after a recycle' {
        $image = New-Object System.Windows.Controls.Image
        Set-GridRowImage -Image $image -Item ([pscustomobject]@{ Script = 'package-office.ps1' })
        Set-GridRowImage -Image $image -Item ([pscustomobject]@{ Script = 'package-nothere.ps1' })
        $image.Source | Should -BeOfType ([System.Windows.Media.DrawingImage])
    }

    It 'leaves an image alone when the item is no packager row' {
        $image = New-Object System.Windows.Controls.Image
        Set-GridRowImage -Image $image -Item ([pscustomobject]@{ Script = 'package-office.ps1' })
        $before = $image.Source
        Set-GridRowImage -Image $image -Item ([pscustomobject]@{ Other = 1 })
        [object]::ReferenceEquals($image.Source, $before) | Should -BeTrue
    }
}

Describe 'Grid icon column' {
    BeforeAll {
        $script:Xaml = Get-Content -LiteralPath (Join-Path $script:Root 'MainWindow.xaml') -Raw
        $script:Start = Get-Content -LiteralPath $script:StartPath -Raw
    }

    It 'has an icon column with an image after the select column and before the application' {
        $script:Xaml | Should -Match 'x:Name="colSelected"'
        $script:Xaml | Should -Match 'x:Name="colIcon"'
        $script:Xaml | Should -Match '<Image Width="24"'
        $script:Xaml.IndexOf('x:Name="colSelected"') | Should -BeLessThan $script:Xaml.IndexOf('x:Name="colIcon"')
        $script:Xaml.IndexOf('x:Name="colIcon"') | Should -BeLessThan $script:Xaml.IndexOf('Header="Application"')
    }

    It 'sets the icon on the cell image from the row item, not through a binding' {
        $script:Xaml | Should -Not -Match 'Binding Icon'
        $script:Start | Should -Match 'Add_LoadingRow'
        $script:Start | Should -Match 'Set-GridRowImage -Image \$image -Item \$item'
        $script:Start | Should -Match 'Get-GridIconSource -ScriptName \(\[string\]\$Item\.Script\)'
    }

    It 'never skips an image because it already has a source (a recycled row container keeps its Image)' {
        $script:Start | Should -Not -Match '\$null -eq \$image\.Source'
        $script:Start | Should -Not -Match '\$null -ne \$image\.Source'
        $script:Start | Should -Match 'add_DataContextChanged'
        $script:Start | Should -Match 'Set-GridRowImage -Image \$img -Item \$args2\.NewValue'
    }

    It 'reads the row at the time the icon is set, not the row the handler was created for' {
        $t = $null; $e = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($script:Start, [ref]$t, [ref]$e)
        $handler = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'Add_LoadingRow' }, $true)
        $handler | Should -Not -BeNullOrEmpty
        $handler.Extent.Text | Should -Match '\$item = \$row\.Item'
        $handler.Extent.Text | Should -Not -Match '\$e\.Row\.Item\s*?
\s*if \(-not \$item'
    }

    It 'refreshes the grid icons after an icon pack install' {
        $script:Start | Should -Match 'if \(\$outcome -and \$outcome\.Installed\) \{ Update-GridIcons \}'
    }
}
