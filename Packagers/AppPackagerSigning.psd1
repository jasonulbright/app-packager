@{
    RootModule        = 'AppPackagerSigning.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = '3f2b8c14-6d2a-4e51-9a77-0c5b41d8e6b2'
    Author            = 'AppPackager'
    Description       = 'Authenticode signing service for staged AppPackager content.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Policy and certificate selection
        'Get-SigningPolicy'
        'Get-CodeSigningCertificateCandidates'
        'Resolve-SigningCertificate'
        'Test-SigningConfiguration'

        # Signing and verification
        'Invoke-ScriptSigning'
        'Test-ScriptSignature'
        'Test-ScriptSignatureBytes'
        'Get-SignedScriptRepresentation'
        'Test-SignedScriptRoundTrip'

        # Launchers
        'New-DeploymentLauncherCommand'
        'Test-DeploymentLauncherChain'

        # Orchestration
        'Invoke-CategorySigning'
        'Get-ConfigMgrDetectionScriptMaxBytes'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}
