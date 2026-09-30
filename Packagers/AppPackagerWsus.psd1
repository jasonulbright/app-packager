@{
    RootModule        = 'AppPackagerWsus.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = 'edadf35c-ae15-4ded-b15c-c44b22fe95b5'
    Author            = 'AppPackager'
    Description       = 'Publishes staged AppPackager installers to WSUS as locally published updates.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Settings and identity
        'Get-WsusClassificationNames'
        'ConvertTo-WsusPublishSettings'
        'Get-WsusIdentityTag'
        'New-WsusPackageId'
        'Get-WsusIdentityLine'
        'Get-WsusUpdateIdentity'
        'ConvertFrom-WsusCatalogInput'

        # Mapping
        'ConvertTo-WsusMsiCommandLine'
        'ConvertTo-WsusApplicabilityRules'
        'Get-WsusCompatibilityFindings'

        # Server
        'Test-WsusAdministrationApi'
        'Get-WsusServerStatus'
        'New-WsusSelfSignedSigningCertificate'
        'Set-WsusSigningCertificate'
        'Export-WsusSigningCertificate'
        'Get-WsusComputerGroupNames'

        # Publishing and server content
        'Publish-WsusSoftwareUpdate'
        'Get-WsusPublishedUpdates'
        'Set-WsusPublishedUpdateState'
        'Import-WsusCatalogUpdate'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
