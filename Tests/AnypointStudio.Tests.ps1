BeforeAll {
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-anypointstudio.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Get-AnypointStudioZipUrl', 'Get-AnypointStudioCandidateVersions', 'Get-AnypointStudioFeatureVersion') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }
    $script:DownloadBase = 'https://mule-studio.s3.amazonaws.com'

    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    function New-TestZip {
        param([string]$Path, [string[]]$Entries)
        $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $Entries) {
                $w = New-Object System.IO.StreamWriter($zip.CreateEntry($name).Open())
                $w.Write('x'); $w.Dispose()
            }
        }
        finally { $zip.Dispose() }
    }
}

Describe 'Anypoint Studio download URL' {
    It 'builds the GA win64 ZIP path for a version' {
        Get-AnypointStudioZipUrl -Version '7.25.0' | Should -Be 'https://mule-studio.s3.amazonaws.com/7.25.0-GA/AnypointStudio-7.25.0-win64.zip'
    }
}

Describe 'Anypoint Studio candidate versions' {
    It 'starts one minor above the newest listed version and walks down, patches 3 to 0' {
        $c = @(Get-AnypointStudioCandidateVersions -Newest ([version]'7.28.0'))
        $c[0] | Should -Be '7.29.3'
        $c[4] | Should -Be '7.28.3'
        $c | Should -Contain '7.25.0'
        $c[-1] | Should -Be '7.20.0'
        $c.Count | Should -Be 40
    }

    It 'stops at minor 0' {
        $c = @(Get-AnypointStudioCandidateVersions -Newest ([version]'8.2.0'))
        $c[-1] | Should -Be '8.0.0'
        $c.Count | Should -Be 16
    }
}

Describe 'Anypoint Studio feature version' {
    It 'reads the full build from the Studio feature folder' {
        $zip = Join-Path $TestDrive 'ok.zip'
        New-TestZip -Path $zip -Entries @(
            'AnypointStudio/AnypointStudio.exe',
            'AnypointStudio/features/org.mule.tooling.studio_7.25.0.202605141228/feature.xml',
            'AnypointStudio/features/org.mule.tooling.p2_7.25.0.202605141228/feature.xml'
        )
        Get-AnypointStudioFeatureVersion -ZipPath $zip | Should -Be '7.25.0.202605141228'
    }

    It 'refuses a ZIP without the launcher' {
        $zip = Join-Path $TestDrive 'noexe.zip'
        New-TestZip -Path $zip -Entries @('AnypointStudio/features/org.mule.tooling.studio_7.25.0.1/feature.xml')
        { Get-AnypointStudioFeatureVersion -ZipPath $zip } | Should -Throw '*AnypointStudio.exe*'
    }

    It 'refuses a ZIP without the Studio feature' {
        $zip = Join-Path $TestDrive 'nofeature.zip'
        New-TestZip -Path $zip -Entries @('AnypointStudio/AnypointStudio.exe')
        { Get-AnypointStudioFeatureVersion -ZipPath $zip } | Should -Throw '*org.mule.tooling.studio*'
    }
}
