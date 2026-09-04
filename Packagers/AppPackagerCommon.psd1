@{
    RootModule        = 'AppPackagerCommon.psm1'
    ModuleVersion     = '0.0.22'
    GUID              = 'f5cdd2d6-eb09-47bd-8493-16dfd5666455'
    Author            = 'AppPackager'
    Description       = 'Shared helpers for AppPackager packager scripts.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging

        # Download
        'Invoke-DownloadWithRetry'
        'Get-GitHubApiCurlArgs'

        # Environment / pre-flight
        'Test-IsAdmin'
        'Connect-CMSite'
        'Initialize-Folder'
        'Test-NetworkShareAccess'

        # Network path
        'Get-NetworkAppRoot'
        'Get-NetworkContentPath'

        # MSI / ARP
        'Get-MsiPropertyMap'
        'Find-UninstallEntry'

        # Installer icons
        'Get-InstallerIcon'
        'Get-MsiIconBytes'
        'Get-IconBytesDimension'
        'ConvertTo-SingleLargestIconFrame'
        'Get-PackagerIconSource'
        'Add-StageIcon'
        'Set-CMApplicationIconFromManifest'
        'Get-IconMimeContent'

        # Stage manifest
        'Get-StageFileHashes'
        'Compare-StageFileHashes'
        'Sync-StagedContentToNetwork'
        'Write-StageManifest'
        'Read-StageManifest'

        # Content wrappers
        'Write-ContentWrappers'
        'New-MsiWrapperContent'
        'New-ExeWrapperContent'
        'New-MsixWrapperContent'

        # MECM
        'New-MECMApplicationFromManifest'
        'Remove-CMApplicationRevisionHistoryByCIId'
        'Get-NextPatchVersion'
        'Test-PsadtLayout'

        # Deployment type requirement rules
        'Get-ConditionTemplatesPath'
        'Get-DefaultConditionTemplates'
        'Get-ConditionTemplates'
        'Save-ConditionTemplates'
        'New-VpnConditionScriptText'
        'Get-OrCreateGlobalConditionFromTemplate'
        'Get-DeploymentTypeRequirementSpecs'
        'New-DeploymentTypeRequirementRules'
        'Get-ManifestDeploymentTypeSpecs'
        'Get-RequestedPackagerVariants'
        'Get-RequestedCommandOverrides'

        # Intune Win32 publishing
        'Get-IntuneWinEncryptionInfo'
        'Export-IntuneWinPayload'
        'Get-MsGraphToken'
        'Invoke-GraphJson'
        'Invoke-AzureBlobUpload'
        'ConvertTo-IntuneWin32Rules'
        'Publish-IntuneWin32App'

        # Intune Win32 content prep
        'Install-IntuneWinAppUtil'
        'New-IntuneWinPackage'

        # Preferences
        'Get-PackagerPreferences'

        # ODT config XML
        'New-OdtConfigXml'

        # Java vendor release helpers
        'Get-LatestTemurinRelease'
        'Get-LatestCorrettoRelease'

        # Packager history (per-app timestamps; local profile storage, never in repo)
        'Get-PackagerHistoryPath'
        'Read-PackagerHistory'
        'Save-PackagerHistory'
        'Update-PackagerHistory'

        # Ad-hoc drop intake
        'Get-InstallerAnalysis'
        'New-AdHocStage'
        'Invoke-AdHocPackage'
        'New-PackagerFromDrop'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
