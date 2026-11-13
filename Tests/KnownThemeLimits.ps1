# The known limits of the theme that the app cannot fix itself. Theme.Guard.Tests.ps1 reads this list.
# Each entry needs an Id, the place (Where), the reason that says why the limit stays (Reason) and its owner (Source).

$script:KnownThemeLimit = @(
    @{
        Id     = 'sidebar-disabled-white-on-blue'
        Where  = 'The sidebar buttons of the main window in the light theme'
        Reason = 'SuiteCommon (Set-ThemeForeground) sets a local white foreground on a blue button in the light theme. A local value beats the gray text of the disabled state in the button style, so a disabled sidebar button would keep white text on blue and look enabled. AppPackager disables the sidebar buttons while a run is active (function Set-ActionButtonsEnabled), so the limit shows during a run.'
        Source = 'Lib\SuiteCommon\SuiteCommon.psm1, Set-ThemeForeground. The owner of the library is app-packager-suite. The fix belongs there.'
    }
)
