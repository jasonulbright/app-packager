@{
    RootModule        = 'AppPackagerOneClick.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '5b1d0f2e-6c1a-4d7e-9a3b-2f8c4e1d7a90'
    Author            = 'AppPackager'
    Description       = 'Plans One Click runs per destination, records what each run published, and writes the run report.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        'Get-OneClickDestinationNames'
        'ConvertTo-OneClickDestinationSet'
        'Get-OneClickDestinations'
        'Test-OneClickDestinationInUse'
        'Get-OneClickDestinationScope'
        'Get-OneClickPublishedVersion'
        'Set-OneClickPublishedVersion'
        'Get-OneClickNotSupportedVersion'
        'Get-OneClickNotSupportedReason'
        'Set-OneClickNotSupported'
        'Get-OneClickCadenceDays'
        'Get-OneClickPlan'
        'Get-OneClickPlanSummary'
        'Select-OneClickRunDestinations'
        'Get-OneClickReportFolder'
        'Write-OneClickReport'
        'Get-OneClickReportList'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
