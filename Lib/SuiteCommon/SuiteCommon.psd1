@{
    RootModule        = 'SuiteCommon.psm1'
    ModuleVersion     = '0.3.0'
    GUID              = '7c1f2a9e-4b3d-4f8a-9c6e-2d5b8e1a7f40'
    Author            = 'Jason Ulbright'
    Description       = 'Shared core for the MECM tool suite: logging, CM site connection, settings persistence, window chrome, theming, dialogs.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging
        'Initialize-Logging'
        'Write-Log'
        'Write-LogErrorRecord'

        # CM Connection
        'Resolve-ConfigurationManagerModulePath'
        'Connect-CMSite'
        'Disconnect-CMSite'
        'Test-CMConnection'
        'Get-CMConnectionInfo'
        'Test-CMSiteCodeMatchesProvider'

        # Settings
        'Read-SuiteSettings'
        'Save-SuiteSettings'

        # Identity
        'Get-SuiteCommonVersion'

        # Window chrome
        'Get-TitleBarDragHeight'
        'Get-InputAncestors'
        'Test-IsWindowCommandPoint'
        'Add-NativeTitleBarHitTestHook'
        'Remove-NativeTitleBarHitTestHook'
        'Install-TitleBarDragFallback'
        'Save-WindowState'
        'Restore-WindowState'

        # Theming
        'Initialize-SuiteTheme'
        'Update-TitleBarBrushes'
        'Update-SidebarButtonTheme'
        'Set-ButtonTheme'
        'Set-DialogTheme'

        # Dialogs
        'Show-ThemedMessage'
        'Show-ConfirmDialog'

        # Background work
        'New-SuiteBgRunspace'
        'Stop-SuiteBgWork'
        'Close-SuiteBgRunspace'

        # CM widgets
        'Build-CollectionTree'
        'Show-CollectionPickerDialog'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
