@{
    RootModule        = 'AppPackagerDiscovery.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '55db0b0c-9651-4da6-934c-8d42c7631cbf'
    Author            = 'AppPackager'
    Description       = 'Reads Add or Remove Programs inventory from the ConfigMgr site database and finds hidden packagers that clients already have installed.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        'Resolve-CMDatabaseTarget'
        'Get-CMInventorySnapshot'
        'Find-DiscoveryCandidate'
        'Get-DiscoveryStatusText'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
