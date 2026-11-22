#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The per-application installer type (EXE or MSI): the header tag, the value the packager reads, and what each of the three packagers
    that declare both types stages for each type.

.DESCRIPTION
    Windows PowerShell 5.1 and PowerShell 7. The cases open no window, download nothing and install nothing. The packagers are read
    as text, so that the silent switches, the detection and the wrapper of each type are checked where they are written.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot 'start-apppackager.ps1'), [ref]$tokens, [ref]$errors)
    if ($errors) { throw ($errors.Message -join '; ') }
    foreach ($name in 'Get-PackagerMetadata', 'ConvertTo-InstallerTypeValue') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "start-apppackager.ps1 no longer defines $name." }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    function Get-PackagerText { param([string]$Name) Get-Content -LiteralPath (Join-Path $script:RepoRoot "Packagers\$Name.ps1") -Raw }
    $script:SavedType = $env:APP_PACKAGER_INSTALLER_TYPE
}

AfterAll { $env:APP_PACKAGER_INSTALLER_TYPE = $script:SavedType }

Describe 'The header tag SupportsInstallerTypes' {
    It 'reads EXE and MSI in any case, and drops a word that is neither' {
        $path = Join-Path $TestDrive 'package-both.ps1'
        Set-Content -LiteralPath $path -Encoding ASCII -Value @'
<#
Vendor: Contoso
App: Widget
SupportsInstallerTypes: exe, MSI, pkg
#>
'@
        @((Get-PackagerMetadata -Path $path).SupportsInstallerTypes) | Should -Be @('EXE', 'MSI')
    }
    It 'reads a packager without the tag as having no choice' {
        $path = Join-Path $TestDrive 'package-plain.ps1'
        Set-Content -LiteralPath $path -Encoding ASCII -Value "<#`nVendor: Contoso`nApp: Widget`n#>"
        @((Get-PackagerMetadata -Path $path).SupportsInstallerTypes).Count | Should -Be 0
    }
    It 'is declared by <Name>' -ForEach @(
        @{ Name = 'package-7zip' }, @{ Name = 'package-notepadplusplus' }, @{ Name = 'package-teamviewerhost' }) {
        @((Get-PackagerMetadata -Path (Join-Path $script:RepoRoot "Packagers\$Name.ps1")).SupportsInstallerTypes) | Should -Be @('EXE', 'MSI')
    }
    It 'is declared by these three packagers only' {
        $declaring = @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Packagers') -Filter 'package-*.ps1' -File |
            Where-Object { @(Get-Content -LiteralPath $_.FullName -TotalCount 40) -match '^\s*SupportsInstallerTypes\s*:' } |
            ForEach-Object { $_.BaseName })
        $declaring | Should -Be @('package-7zip', 'package-notepadplusplus', 'package-teamviewerhost')
    }
}

Describe 'The value that the packager reads' {
    It 'maps a stored entry to EXE, MSI or nothing' {
        ConvertTo-InstallerTypeValue -Entry ([pscustomobject]@{ InstallerType = 'MSI' }) | Should -Be 'MSI'
        ConvertTo-InstallerTypeValue -Entry ([pscustomobject]@{ InstallerType = 'EXE' }) | Should -Be 'EXE'
        ConvertTo-InstallerTypeValue -Entry ([pscustomobject]@{ InstallerType = 'zip' }) | Should -Be ''
        ConvertTo-InstallerTypeValue -Entry ([pscustomobject]@{ InstallMode = 'AllUsers' }) | Should -Be ''
        ConvertTo-InstallerTypeValue -Entry $null | Should -Be ''
    }
    It 'reads APP_PACKAGER_INSTALLER_TYPE and refuses another value' {
        $env:APP_PACKAGER_INSTALLER_TYPE = $null
        Get-RequestedInstallerType | Should -BeNullOrEmpty
        $env:APP_PACKAGER_INSTALLER_TYPE = ' msi '
        Get-RequestedInstallerType | Should -Be 'MSI'
        $env:APP_PACKAGER_INSTALLER_TYPE = 'zip'
        { Get-RequestedInstallerType } | Should -Throw '*must be EXE or MSI*'
    }
}

Describe 'package-7zip' {
    BeforeAll { $script:Text = Get-PackagerText 'package-7zip' }
    It 'stages the MSI by default with msiexec /qn and the ProductCode key' {
        $script:Text | Should -Match ([regex]::Escape("if (-not `$installerType) { `$installerType = 'MSI' }"))
        $script:Text | Should -Match ([regex]::Escape('New-MsiWrapperContent -MsiFileName $MsiFileName'))
        $script:Text | Should -Match ([regex]::Escape('"/qn /norestart"'))
        $script:Text | Should -Match ([regex]::Escape('Uninstall\" + $productCode'))
    }
    It 'stages the x64 exe on request with /S, Uninstall.exe /S and the fixed 7-Zip key' {
        $script:Text | Should -Match ([regex]::Escape('Resolve-7ZipArm64ExeUrl -Architecture x64'))
        $script:Text | Should -Match ([regex]::Escape('New-ExeWrapperContent -InstallerFileName $ExeFileName -InstallArgs "''/S''"'))
        $script:Text | Should -Match ([regex]::Escape('C:\Program Files\7-Zip\Uninstall.exe'))
        $script:Text | Should -Match ([regex]::Escape('CurrentVersion\Uninstall\7-Zip'''))
        $script:Text | Should -Match ([regex]::Escape("ProductCode'] = `$productCode"))
    }
    It 'refuses an exe whose version differs from the MSI' {
        $script:Text | Should -Match 'does not match the MSI version'
    }
}

Describe 'package-notepadplusplus' {
    BeforeAll { $script:Text = Get-PackagerText 'package-notepadplusplus' }
    It 'keeps the NSIS exe as the default with /S /noUpdater' {
        $script:Text | Should -Match ([regex]::Escape("if (-not `$installerType) { `$installerType = 'EXE' }"))
        $script:Text | Should -Match ([regex]::Escape("ArgumentList @(''/S'', ''/noUpdater'')"))
        $script:Text | Should -Match 'uninstall\.exe'
    }
    It 'takes the x64 MSI asset of the same release for the MSI type, with /qn' {
        $script:Text | Should -Match ([regex]::Escape("'*Installer.x64.msi'"))
        $script:Text | Should -Match ([regex]::Escape('New-MsiWrapperContent -MsiFileName $installerFileName'))
        $script:Text | Should -Match ([regex]::Escape('"/qn /norestart"'))
        $script:Text | Should -Match ([regex]::Escape("ProductCode'] = `$msiProductCode"))
    }
    It 'detects the same file version for both types' {
        ([regex]::Matches($script:Text, [regex]::Escape('FileName      = "notepad++.exe"'))).Count | Should -Be 1
    }
}

Describe 'package-teamviewerhost' {
    BeforeAll { $script:Text = Get-PackagerText 'package-teamviewerhost' }
    It 'keeps the user-provided MSI as the default' {
        $script:Text | Should -Match ([regex]::Escape("if (-not `$type) { return 'MSI' }"))
        $script:Text | Should -Match 'Resolve-LocalSourceInstaller'
        $script:Text | Should -Match ([regex]::Escape('InstallerType   = "MSI"'))
    }
    It 'stages the vendor exe for the EXE type with /S and a file-version detection' {
        $script:Text | Should -Match 'function Invoke-StageTeamViewerHostExe'
        $script:Text | Should -Match 'TeamViewer_Host_Setup_x64\.exe'
        $script:Text | Should -Match ([regex]::Escape("ArgumentList @(''/S'')"))
        $script:Text | Should -Match ([regex]::Escape('InstallerType   = "EXE"'))
        $script:Text | Should -Match ([regex]::Escape('FileName      = "TeamViewer.exe"'))
    }
    It 'warns that the EXE type ignores the MSI properties of the config file' {
        $script:Text | Should -Match 'MSI properties that the EXE installer type does not pass'
    }
    It 'says in its help that TeamViewer documents the silent install for the MSI only' {
        $script:Text | Should -Match 'does not\s+document an EXE switch'
    }
}

Describe 'Confirm-LocalSourceFolders' {
    It 'asks no folder of an application set to the EXE type, and still asks for the MSI type' {
        $text = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'start-apppackager.ps1') -Raw
        $text | Should -Match ([regex]::Escape("SupportsInstallerTypes) -contains 'EXE'"))
        $text | Should -Match ([regex]::Escape("if (`$chosen -eq 'EXE') { `$kept.Add(`$row); continue }"))
    }
}
