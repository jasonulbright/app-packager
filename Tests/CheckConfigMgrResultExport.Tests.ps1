#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The result file of Check ConfigMgr: the rows, the CSV, the JSON twin, the filter and the context menu entry.

.DESCRIPTION
    The format is the same as the one App Packager Desktop writes (CheckConfigMgrResult.csv and CheckConfigMgrResult.json).
    The Update rollout view of Deployment Helper reads it. A fake grid list; no window and no ConfigMgr.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    Import-Module (Join-Path $script:RepoRoot 'Packagers\AppPackagerCommon.psd1') -Force -DisableNameChecking
    $script:At = [datetime]::SpecifyKind([datetime]'2026-11-21 08:30:05', 'Utc')

    function New-Row($Application, $Status, $Cm = '1.0', $Latest = '2.0', $Arp = '', $Arch = '') {
        [pscustomobject]@{ Application = $Application; Vendor = 'Vendor, Inc'; CMName = "CM $Application"; ArpName = $Arp; ArpArch = $Arch; CurrentVersion = $Cm; LatestVersion = $Latest; Status = $Status }
    }
}

Describe 'ConvertTo-CheckConfigMgrResultRow' {
    It 'maps <Grid> to <File>' -ForEach @(
        @{ Grid = 'Update available'; File = 'UpdateAvailable' }
        @{ Grid = 'Update available (3 CVEs, 1 in KEV)'; File = 'UpdateAvailable' }
        @{ Grid = 'Up to date'; File = 'UpToDate' }
        @{ Grid = 'Current newer'; File = 'UpToDate' }
        @{ Grid = 'Not found in ConfigMgr'; File = 'NotFound' }
    ) {
        @(ConvertTo-CheckConfigMgrResultRow -Rows (New-Row 'A' $Grid) -CheckedAtUtc $script:At)[0].Status | Should -Be $File
    }

    It 'leaves out a row without a usable answer' {
        $rows = 'Error', 'Cancelled', 'ConfigMgr version retrieved', 'Ready', '' | ForEach-Object { New-Row 'A' $_ }
        @(ConvertTo-CheckConfigMgrResultRow -Rows $rows).Count | Should -Be 0
    }

    It 'fills the columns: the ARP names, the architecture, the source and the time' {
        $r = @(ConvertTo-CheckConfigMgrResultRow -Rows (New-Row 'Tool' 'Update available' -Arp 'Tool* | none | Tool Suite*' -Arch 'X86') -CheckedAtUtc $script:At)[0]
        $r.InstalledSoftwareName | Should -Be 'Tool* | Tool Suite*'
        $r.Architecture | Should -Be 'x86'
        $r.Source | Should -Be 'AppPackager'
        $r.CheckedAtUtc | Should -Be '2026-11-21T08:30:05Z'
        $r.ConfigMgrVersion | Should -Be '1.0'
        $r.LatestVersion | Should -Be '2.0'
    }

    It 'takes the architecture from the application name, else x64' {
        @(ConvertTo-CheckConfigMgrResultRow -Rows (New-Row 'Tool (ARM64)' 'Up to date'))[0].Architecture | Should -Be 'arm64'
        @(ConvertTo-CheckConfigMgrResultRow -Rows (New-Row 'Tool' 'Up to date'))[0].Architecture | Should -Be 'x64'
    }
}

Describe 'Export-CheckConfigMgrResult' {
    BeforeEach {
        $script:Dir = Join-Path ([System.IO.Path]::GetTempPath()) ('ccr-' + [guid]::NewGuid().ToString('N'))
        $script:Result = @(ConvertTo-CheckConfigMgrResultRow -CheckedAtUtc $script:At -Rows @(
                (New-Row 'A "x"' 'Update available' -Arp 'A* | B*'), (New-Row 'B' 'Up to date'), (New-Row 'C' 'Not found in ConfigMgr' -Cm '')))
    }
    AfterEach { Remove-Item -LiteralPath $script:Dir -Recurse -Force -ErrorAction SilentlyContinue }

    It 'writes a CSV with a byte order mark, the header and every field in quotes' {
        $p = Join-Path $script:Dir 'r.csv'
        (Export-CheckConfigMgrResult -Result $script:Result -Path $p).Count | Should -Be 3
        $bytes = [System.IO.File]::ReadAllBytes($p)
        $bytes[0..2] | Should -Be @(0xEF, 0xBB, 0xBF)
        $lines = [System.IO.File]::ReadAllText($p).TrimStart([char]0xFEFF).Split(@("`r`n"), [System.StringSplitOptions]::RemoveEmptyEntries)
        $lines[0] | Should -Be '"Application","Vendor","CMName","InstalledSoftwareName","Architecture","ConfigMgrVersion","LatestVersion","Status","Source","CheckedAtUtc"'
        $lines[1] | Should -Be '"A ""x""","Vendor, Inc","CM A ""x""","A* | B*","x64","1.0","2.0","UpdateAvailable","AppPackager","2026-11-21T08:30:05Z"'
        $lines.Count | Should -Be 4
    }

    It 'keeps only the updates with -OnlyUpdatesAvailable' {
        $p = Join-Path $script:Dir 'r.csv'
        (Export-CheckConfigMgrResult -Result $script:Result -Path $p -OnlyUpdatesAvailable).Count | Should -Be 1
    }

    It 'writes the JSON twin with the names of the CSV columns, as an array even for one row' {
        $p = Join-Path $script:Dir 'r.json'
        [void](Export-CheckConfigMgrResult -Result $script:Result -Path $p -OnlyUpdatesAvailable)
        $text = [System.IO.File]::ReadAllText($p).TrimStart([char]0xFEFF)
        $text.StartsWith('[') | Should -BeTrue
        $json = @(($text | ConvertFrom-Json) | ForEach-Object { $_ })
        $json.Count | Should -Be 1
        ($json[0].PSObject.Properties.Name -join ',') | Should -Be 'Application,Vendor,CMName,InstalledSoftwareName,Architecture,ConfigMgrVersion,LatestVersion,Status,Source,CheckedAtUtc'
    }

    It 'writes the fixed pair of files and overwrites them on the next run' {
        [void](Write-CheckConfigMgrResultDefault -Result $script:Result -Folder $script:Dir)
        [void](Write-CheckConfigMgrResultDefault -Result @($script:Result[2]) -Folder $script:Dir)
        @(Get-Content -LiteralPath (Join-Path $script:Dir 'CheckConfigMgrResult.csv')).Count | Should -Be 2
        Test-Path -LiteralPath (Join-Path $script:Dir 'CheckConfigMgrResult.json') | Should -BeTrue
        @(Get-ChildItem -LiteralPath $script:Dir -Filter '*.tmp').Count | Should -Be 0
    }

    It 'uses the AppPackager folder of the local app data' {
        Get-CheckConfigMgrResultFolder | Should -Be (Join-Path $env:LOCALAPPDATA 'AppPackager')
    }
}

Describe 'The window' {
    It 'writes the file after the run and offers the export in the grid context menu, without a new control in the main window' {
        $source = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'start-apppackager.ps1') -Raw
        $source | Should -Match 'Save-CheckConfigMgrResult -Rows @\(\$selectedRows\)'
        $source | Should -Match '\$contextMenu\.Items\.Add\(\$menuExportCheckResult\)'
        (Get-Content -LiteralPath (Join-Path $script:RepoRoot 'MainWindow.xaml') -Raw) | Should -Not -Match 'CheckConfigMgrResult|ExportCheck'
    }
}
