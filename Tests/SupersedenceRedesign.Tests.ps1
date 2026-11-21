#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
<#
.SYNOPSIS
    The supersedence redesign (Q-92) as the window, the Options page and the command line carry it: no button, no menu entry
    and no function of the removed features; the confirmation before a Package run; the title lock.

.DESCRIPTION
    The window is read as source, as the other shell tests do.
#>

BeforeAll {
    $script:Root = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:Shell = Get-Content -LiteralPath (Join-Path $script:Root 'start-apppackager.ps1') -Raw
    $script:Xaml = Get-Content -LiteralPath (Join-Path $script:Root 'MainWindow.xaml') -Raw
    $script:Build = Get-Content -LiteralPath (Join-Path $script:Root 'Invoke-AppPackagerBuild.ps1') -Raw
}

Describe 'What was removed' {
    It 'has no sidebar button for Apply Supersedence, Rewrite Old Versions or Restore Old Versions' {
        $script:Xaml | Should -Not -Match 'btnApplySupersedence|btnRewriteOld|btnRestoreOld'
        $script:Shell | Should -Not -Match 'btnApplySupersedence|btnRewriteOld|btnRestoreOld'
    }
    It 'has no old-application window and no menu entry for it' {
        $script:Shell | Should -Not -Match 'Use as old application|Open-OldApplicationDialog|Show-OldApplicationDialog'
    }
    It 'has none of the functions of the rewrite, the restore, the apply or the One Click pass' {
        foreach ($name in 'Start-RewriteOldVersionsRun', 'Start-ApplySupersedence', 'Show-RewriteOldVersionsDialog', 'Invoke-RewriteChild', 'Invoke-OneClickSupersedencePass') {
            $script:Shell | Should -Not -Match ('function ' + [regex]::Escape($name))
        }
    }
    It 'has no command line option of the rewrite' {
        $script:Build | Should -Not -Match 'RewriteOldVersions|IdentityMarker|ShownPlan'
    }
}

Describe 'The confirmation before a Package run' {
    It 'asks before the package child runs, with Yes, Skip and No, and keeps Remember for the run' {
        $script:Shell | Should -Match 'function Show-SupersedenceConfirmDialog'
        $script:Shell | Should -Match 'Content="Yes"'
        $script:Shell | Should -Match 'Content="Skip"'
        $script:Shell | Should -Match 'Content="No"'
        $script:Shell | Should -Match 'SupersedenceDecisionForAll'
        $script:Shell | Should -Match 'Get-SupersedenceRunAnswer -State \$State'
    }
    It 'stops the whole run on No' {
        $script:Shell | Should -Match 'if \(\$supersede\.Cancel\) \{\s+\$State\.CancelRequested = \$true'
    }
    It 'passes the answer to the child, and Skip when nobody was asked' {
        $script:Shell | Should -Match "APP_PACKAGER_SUPERSEDENCE_ANSWER'\] = \`$\(if \(\`$SupersedenceAnswer -eq 'Yes'\)"
        $script:Shell | Should -Match "\[ValidateSet\('', 'Yes', 'Skip'\)\]\[string\]\`$SupersedenceAnswer"
    }
    It 'locks the version in the title for an application with supersedence on, at every Package path' {
        ([regex]::Matches($script:Shell, 'Get-SupersedenceLockedTitleMode -PackagerPath')).Count | Should -BeGreaterOrEqual 3
    }
}

Describe 'The Options page' {
    It 'has the setting by application and the bulk action next to the default' {
        $script:Shell | Should -Match "Id = 'sup.apps'"
        $script:Shell | Should -Match "'All applications'"
        $script:Shell | Should -Match "'Checked applications'"
        $script:Shell | Should -Match "'No applications'"
    }
    It 'keeps the default Off' {
        $script:Shell | Should -Match "Pref = 'Supersedence.Mode'; Default = 'Off'"
    }
    It 'has the note field of the title page, and no code that writes the include-version box from the default' {
        # The behaviour is tested in Invoke-SupersedenceOptionsSmoke.ps1 (window probe) and SupersedenceStep.Tests.ps1 (the note text).
        $script:Shell | Should -Match "Id = 'title.lock'"
        $script:Shell | Should -Not -Match 'includeBox2'
    }
}
