BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1') -Force
    $t = $null; $e = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\Packagers\package-anypointstudio.ps1'), [ref]$t, [ref]$e)
    if ($e) { throw ($e.Message -join '; ') }
    foreach ($name in 'Get-AnypointStudioReleaseFromManifest', 'Get-AnypointStudioFeatureVersion', 'Assert-AnypointStudioZipLayout', 'New-AnypointStudioInstallContent') {
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $false)
        if (-not $fn) { throw "function $name not found" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

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

Describe 'Anypoint Studio downloads manifest' {
    BeforeAll {
        $script:manifest = @'
[
  { "id": 0, "name": "studio", "version": "latest", "os": "linux", "source": "https://www.mulesoft.com/downloads/studio/latest/AnypointStudio-7.28.0-linux64.tar.gz", "integrity": "3a97a7da3cfd1f354b2db069f034f0e6bae10e09c47e64e82e4d06cb66f71669", "packaging": "tar" },
  { "id": 1, "name": "studio", "version": "latest", "os": "windows", "source": "https://www.mulesoft.com/downloads/studio/latest/AnypointStudio-7.28.0-win64.zip", "integrity": "bec2ef0a0d1c0f60276e9bb0683d7131ad8cb5bb9b277c5bfa94ec43115f40ac", "packaging": "zip" },
  { "id": 4, "name": "studio", "version": "previous", "os": "windows", "source": "https://www.mulesoft.com/downloads/studio/previous/AnypointStudio-7.20.1-win64.zip", "integrity": "cd09b75448183d1a2f755b6dc38d0a4fd91a9d922b4a57aa75b53f5acd95f062", "packaging": "zip" },
  { "id": 6, "name": "mule", "version": "latest", "source": "https://www.mulesoft.com/downloads/mule/latest/mule-ee-distribution-standalone-4.12.3.zip", "integrity": "ace17d7b2313abb9e386b2b2695f475a78145a2168ad0564c2cfa6ccf500cd2b", "packaging": "zip" }
]
'@
    }

    It 'picks the latest Windows Studio entry with its version and SHA-256' {
        $r = Get-AnypointStudioReleaseFromManifest -Json $script:manifest
        $r.Version | Should -Be '7.28.0'
        $r.FileName | Should -Be 'AnypointStudio-7.28.0-win64.zip'
        $r.DownloadUrl | Should -Be 'https://www.mulesoft.com/downloads/studio/latest/AnypointStudio-7.28.0-win64.zip'
        $r.Sha256 | Should -Be 'BEC2EF0A0D1C0F60276E9BB0683D7131AD8CB5BB9B277C5BFA94EC43115F40AC'
    }

    It 'refuses a manifest without a latest Windows Studio entry' {
        $json = '[{ "name": "studio", "version": "previous", "os": "windows", "source": "https://x/AnypointStudio-7.20.1-win64.zip", "integrity": "00" }]'
        { Get-AnypointStudioReleaseFromManifest -Json $json } | Should -Throw '*no latest Windows Studio entry*'
    }

    It 'refuses an unexpected file name' {
        $json = '[{ "name": "studio", "version": "latest", "os": "windows", "source": "https://x/AnypointStudio-latest.zip", "integrity": "bec2ef0a0d1c0f60276e9bb0683d7131ad8cb5bb9b277c5bfa94ec43115f40ac" }]'
        { Get-AnypointStudioReleaseFromManifest -Json $json } | Should -Throw '*Unexpected Windows Studio file name*'
    }

    It 'refuses an entry without a SHA-256' {
        $json = '[{ "name": "studio", "version": "latest", "os": "windows", "source": "https://x/AnypointStudio-7.28.0-win64.zip", "integrity": "" }]'
        { Get-AnypointStudioReleaseFromManifest -Json $json } | Should -Throw '*no SHA-256*'
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

Describe 'Anypoint Studio ZIP layout' {
    BeforeAll {
        $script:Parent = Join-Path $TestDrive 'layout'
        $script:Install = Join-Path $script:Parent 'AnypointStudio'
    }

    It 'accepts entries that stay under the install folder' {
        $zip = Join-Path $TestDrive 'layout-ok.zip'
        New-TestZip -Path $zip -Entries @('AnypointStudio/', 'AnypointStudio/AnypointStudio.exe', 'AnypointStudio/plugins/a/b.jar', 'AnypointStudio/x/../y.txt')
        { Assert-AnypointStudioZipLayout -ZipPath $zip -InstallDir $script:Install } | Should -Not -Throw
    }

    It 'refuses the entry <Entry>' -ForEach @(
        @{ Entry = 'AnypointStudio/../escaped.txt' }
        @{ Entry = 'AnypointStudio/../../escaped.txt' }
        @{ Entry = 'AnypointStudio/a/../../escaped.txt' }
        @{ Entry = 'Other/file.txt' }
        @{ Entry = 'AnypointStudioEvil/file.txt' }
        @{ Entry = 'C:/escaped.txt' }
        @{ Entry = '/escaped.txt' }
    ) {
        $zip = Join-Path $TestDrive ('layout-bad-' + [guid]::NewGuid().ToString('N') + '.zip')
        New-TestZip -Path $zip -Entries @('AnypointStudio/AnypointStudio.exe', $Entry)
        { Assert-AnypointStudioZipLayout -ZipPath $zip -InstallDir $script:Install } | Should -Throw '*extracts outside*'
    }
}

Describe 'Anypoint Studio install script' {
    BeforeAll {
        function script:Invoke-InstallScript {
            param([string]$Name, [string[]]$Entries, [switch]$ExistingInstall)
            $root = Join-Path $TestDrive $Name
            $content = Join-Path $root 'content'
            $installDir = Join-Path $root 'target\AnypointStudio'
            New-Item -ItemType Directory -Path $content -Force | Out-Null
            if ($ExistingInstall) {
                New-Item -ItemType Directory -Path $installDir -Force | Out-Null
                Set-Content -LiteralPath (Join-Path $installDir 'old.txt') -Value 'old' -Encoding ASCII
            }
            New-TestZip -Path (Join-Path $content 'studio.zip') -Entries $Entries
            $script = New-AnypointStudioInstallContent -ZipFileName 'studio.zip' -InstallDir $installDir
            Set-Content -LiteralPath (Join-Path $content 'install.ps1') -Value $script -Encoding ASCII
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $content 'install.ps1') *> $null
            [pscustomobject]@{ Exit = $LASTEXITCODE; Root = $root; InstallDir = $installDir }
        }
    }

    It 'parses without errors' {
        $tokens = $null; $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseInput((New-AnypointStudioInstallContent -ZipFileName 'a.zip' -InstallDir 'C:\AnypointStudio'), [ref]$tokens, [ref]$errors)
        $errors | Should -BeNullOrEmpty
    }

    It 'extracts a ZIP that stays under the install folder and replaces an earlier version' {
        $r = Invoke-InstallScript -Name 'install-ok' -Entries @('AnypointStudio/AnypointStudio.exe', 'AnypointStudio/plugins/p.jar') -ExistingInstall
        $r.Exit | Should -Be 0
        Test-Path -LiteralPath (Join-Path $r.InstallDir 'AnypointStudio.exe') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $r.InstallDir 'old.txt') | Should -BeFalse
    }

    It 'writes nothing and keeps the earlier version for a ZIP with an entry that leaves the install folder' {
        $r = Invoke-InstallScript -Name 'install-bad' -Entries @('AnypointStudio/AnypointStudio.exe', 'AnypointStudio/../escaped.txt') -ExistingInstall
        $r.Exit | Should -Be 1
        Test-Path -LiteralPath (Join-Path $r.Root 'target\escaped.txt') | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $r.InstallDir 'old.txt') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $r.InstallDir 'AnypointStudio.exe') | Should -BeFalse
    }
}
