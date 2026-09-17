BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\start-apppackager.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-PackagerFolderInfo' }, $false)
    . ([scriptblock]::Create($fn.Extent.Text))

    function New-FolderInfoScript {
        param([Parameter(Mandatory)][string]$Body)
        $path = Join-Path $TestDrive ('package-' + [guid]::NewGuid().ToString('N') + '.ps1')
        Set-Content -LiteralPath $path -Value $Body -Encoding ASCII
        return $path
    }
}

Describe 'Packager download subfolder' {
    It 'reads a double-quoted subfolder' {
        $path = New-FolderInfoScript -Body @'
$VendorFolder = "Contoso"
$AppFolder    = "Widget"
$BaseDownloadRoot = Join-Path $DownloadRoot "Widget"
'@
        (Get-PackagerFolderInfo -ScriptPath $path).DownloadSubfolder | Should -Be 'Widget'
    }

    It 'reads a single-quoted subfolder' {
        $path = New-FolderInfoScript -Body @'
$VendorFolder     = 'Contoso'
$AppFolder        = 'Spectra PDF'
$BaseDownloadRoot = Join-Path $DownloadRoot 'SpectraPDF'
'@
        $info = Get-PackagerFolderInfo -ScriptPath $path
        $info.DownloadSubfolder | Should -Be 'SpectraPDF'
        $info.AppFolder | Should -Be 'Spectra PDF'
        $info.VendorFolder | Should -Be 'Contoso'
    }

    It 'resolves a subfolder named through the application folder variable' {
        $path = New-FolderInfoScript -Body @'
$VendorFolder = "Scooter Software"
$AppFolder    = "Beyond Compare 5"
$BaseDownloadRoot = Join-Path $DownloadRoot $AppFolder
'@
        (Get-PackagerFolderInfo -ScriptPath $path).DownloadSubfolder | Should -Be 'Beyond Compare 5'
    }

    It 'resolves a subfolder declared before the folder variables' {
        $path = New-FolderInfoScript -Body @'
$BaseDownloadRoot = Join-Path $DownloadRoot $VendorFolder
$VendorFolder = "Scooter Software"
$AppFolder    = "Beyond Compare 5"
'@
        (Get-PackagerFolderInfo -ScriptPath $path).DownloadSubfolder | Should -Be 'Scooter Software'
    }

    It 'resolves a subfolder for <Name>' -TestCases @(
        foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '..\Packagers') -Filter 'package-*.ps1') +
                          @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot '..\Samples') -Filter 'package-template-*.ps1')) {
            @{ Name = $file.Name; Path = $file.FullName }
        }
    ) {
        param($Name, $Path)
        (Get-PackagerFolderInfo -ScriptPath $Path).DownloadSubfolder | Should -Not -BeNullOrEmpty
    }
}
