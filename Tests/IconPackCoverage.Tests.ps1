#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    Every packager has an icon in the icon pack (owner order D-34: every new packager needs an icon).

.DESCRIPTION
    Tests\IconPackManifest.json is the manifest.json of the pinned pack release (jasonulbright/app-packager-icons, v2026.11.19.0017).
    The test compares the packager ids (the file names Packagers\package-<id>.ps1) with the Packager field of the manifest. A
    packager with no icon in the pack fails the test unless it is on the named list below. The list is empty. A new packager needs
    its icon in a new pack release first: copy the manifest of that release over Tests\IconPackManifest.json and move the pin.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Pin = 'v2026.11.19.0017'
    # Packagers that the pinned pack has no icon for, with the reason for each. A generic installer icon is not shipped instead.
    $script:WithoutIcon = @{}
    $script:Manifest = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'IconPackManifest.json') -Raw | ConvertFrom-Json
    $script:Ids = @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Packagers') -Filter 'package-*.ps1' -File |
            ForEach-Object { $_.BaseName -replace '^package-', '' } | Sort-Object)
    $script:IconIds = @($script:Manifest.Icons | ForEach-Object { ([string]$_.Packager).ToLowerInvariant() })
}

Describe 'Icon pack coverage' {
    It 'pins a manifest of the pack release' {
        $script:Pin | Should -Be ('v' + $script:Manifest.PackVersion)
        @($script:Manifest.Icons).Count | Should -BeGreaterThan 300
    }

    It 'finds the packagers' {
        $script:Ids.Count | Should -BeGreaterThan 300
    }

    It 'has an icon in the pinned pack for every packager, or the packager is on the named list' {
        $missing = @($script:Ids | Where-Object { $script:IconIds -notcontains $_.ToLowerInvariant() })
        $unexpected = @($missing | Where-Object { -not $script:WithoutIcon.ContainsKey($_) })
        $unexpected -join ', ' | Should -BeNullOrEmpty -Because ('these packagers have no icon in the pack ' + $script:Pin + ': ' + ($unexpected -join ', '))
        $fixed = @($script:WithoutIcon.Keys | Where-Object { $_ -notin $missing })
        $fixed -join ', ' | Should -BeNullOrEmpty -Because 'these packagers are on the list but the pack has an icon now: remove them from the list'
    }

    It 'keeps the named list empty' {
        $script:WithoutIcon.Count | Should -Be 0
    }

    It 'names the icon file of every manifest entry after its packager' {
        $bad = @($script:Manifest.Icons | Where-Object { ([string]$_.File) -notmatch ('^' + [regex]::Escape([string]$_.Packager) + '\.(png|ico)$') })
        ($bad | ForEach-Object { $_.File }) -join ', ' | Should -BeNullOrEmpty
    }
}