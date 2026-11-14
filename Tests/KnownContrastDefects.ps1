# Contrast defects of the Options window that exist today and that the app cannot fix itself.
# Each entry is "<control>:<metric>" or "text:<visible text>". Invoke-OptionsContrastSmoke.ps1 fails when a defect
# is not listed (a new defect) and when a listed one no longer happens (remove it here).
#
# titlebar:inactive: the title text on the title bar of a window that is not active. The brush comes from SuiteCommon
# (Update-TitleBarBrushes and Set-DialogTheme: #4BA3E0 in the light theme, the MahApps #808080 in the dark theme).
# The owner of the fix is app-packager-suite. Remove the entries after the vendored copy has the fix.
function Get-KnownContrastDefect {
    param([Parameter(Mandatory)][ValidateSet('dark', 'light')][string]$Theme)
    @('titlebar:inactive')
}
