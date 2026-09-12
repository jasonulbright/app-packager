@{
    RootModule        = 'AppPackagerWorkbench.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = 'a4d17f92-6c3b-4f0a-9b27-1d5e8c3f7a10'
    Author            = 'AppPackager'
    Description       = 'Application Workbench definition, profile, run snapshot and build model.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Storage and identity
        'Get-WorkbenchDataRoot'
        'Get-SharedDownloadCacheRoot'
        'New-ApplicationId'
        'ConvertTo-ApplicationKey'
        'ConvertFrom-ApplicationKey'

        # Applications
        'Get-WorkbenchApplications'
        'Get-ApplicationDefinition'
        'Save-ApplicationDefinition'

        # Profiles
        'New-WorkbenchProfileObject'
        'Get-Profiles'
        'Get-Profile'
        'Save-Profile'
        'Copy-Profile'
        'Remove-ProfileField'
        'Add-ProfileAsset'
        'Remove-ProfileAsset'
        'Resolve-ProfileAssetPath'
        'Resolve-EffectiveSettings'
        'Invoke-LegacyPreferenceMigration'

        # Runs and finalization
        'New-BuildId'
        'New-RunSnapshot'
        'Get-RunSnapshot'
        'Resolve-WorkbenchTokens'
        'New-ExtendOrchestratorContent'
        'Invoke-StageFinalization'

        # Builds
        'Write-BuildRecord'
        'Get-BuildRecords'
        'Get-LatestBuildRecord'
        'Resolve-StageManifestForBuild'

        # BYO
        'Save-ByoApplication'
        'Update-ByoSource'

        # Drafts, bundles, publications
        'Save-Draft'
        'Get-Draft'
        'Remove-Draft'
        'Export-WorkbenchBundle'
        'Import-WorkbenchBundle'
        'Write-PublicationRecord'
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
