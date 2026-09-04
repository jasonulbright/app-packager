@{
    RootModule        = 'InstallerAnalysisCommon.psm1'
    ModuleVersion     = '1.3.0.1'
    GUID              = 'c3d4e5f6-a7b8-9012-cdef-456789012345'
    Author            = 'Jason Ulbright'
    Description       = 'Installer analysis tool - version intelligence, installer type detection, silent switches, payload extraction.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging comes from the vendored SuiteCommon module
        # (Lib\SuiteCommon), imported globally by the root module.

        # File Identification
        'Get-InstallerFileInfo'
        'Get-InstallerType'
        'Get-PeArchitecture'

        # ZIP helpers (used by package-format analyzers)
        'Test-IsZipFile'
        'Test-ZipEntryExists'
        'Get-ZipEntryText'
        'Get-ZipRootEntryByPattern'

        # MSI / MSP Analysis
        'Get-MsiProperties'
        'Get-MsiSummaryInfo'
        'Get-MspMetadata'
        'Test-MsiModuleAvailable'

        # Package-format analyzers (v1.1.0)
        'Get-ChocolateyMetadata'
        'Get-IntunewinMetadata'
        'Get-MsixManifest'
        'ConvertFrom-MsixPackageManifest'
        'ConvertFrom-MsixBundleManifest'
        'Get-PsadtMetadata'
        'ConvertFrom-PsadtDeployApplication'
        'Get-ZipEntryPathByPattern'
        'Get-SquirrelMetadata'
        'Get-WixBurnMetadata'
        'Get-NsisMetadata'
        'Get-PeRequestedExecutionLevel'
        'ConvertTo-NsisWindowsPath'
        'Get-PackageMetadataFor'

        # Deployment Fields
        'Get-DeploymentFields'
        'Get-UninstallRegistryKey'

        # Silent Switches
        'Get-SilentSwitchDatabase'
        'Get-SilentSwitches'

        # Payload Extraction
        'Find-7ZipPath'
        'Get-PayloadContents'
        'Get-EmbeddedInstallers'
        'Expand-PayloadEntry'
        'Expand-InstallerPayload'

        # String Analysis
        'Get-BinaryStrings'
        'Get-InterestingStrings'

        # Export
        'Export-AnalysisReport'
        'Export-AnalysisHtml'
        'New-AnalysisSummaryText'
        'ConvertTo-DeploymentJson'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
