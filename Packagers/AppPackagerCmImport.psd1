@{
    RootModule        = 'AppPackagerCmImport.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '7c3a91e4-5b28-4f6d-a1c7-3e9d20b84f15'
    Author            = 'AppPackager'
    Description       = 'Reads an existing ConfigMgr application and plans, packages and publishes it as Intune Win32 apps, one app per deployment type.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        'ConvertFrom-CMApplicationDigest'
        'ConvertTo-IntuneDetectionRules'
        'New-IntuneExpressionDetectionScript'
        'ConvertTo-IntuneRequirementModel'
        'New-IntuneVariantRequirementScript'
        'Get-CMImportVariantExclusion'
        'ConvertTo-IntuneReturnCodes'
        'Get-ImpSetupFile'
        'New-CMImportPlan'
        'Format-CMImportReport'
        'Get-CMApplicationDigest'
        'Export-CMImportStage'
        'Convert-CMApplicationToIntune'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
