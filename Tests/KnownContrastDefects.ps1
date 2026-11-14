# Contrast defects of the Options window that exist today and that the app cannot fix itself.
# Each entry is "<control>:<metric>" or "text:<visible text>". Invoke-OptionsContrastSmoke.ps1 fails when a defect
# is not listed (a new defect) and when a listed one no longer happens (remove it here).
# The list is empty: the last entry, titlebar:inactive, went away with the SuiteCommon fix.
function Get-KnownContrastDefect {
    param([Parameter(Mandatory)][ValidateSet('dark', 'light')][string]$Theme)
    @()
}
