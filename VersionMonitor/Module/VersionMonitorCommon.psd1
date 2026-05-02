@{
    RootModule        = 'VersionMonitorCommon.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'd4e5f6a7-b8c9-0123-def4-567890123456'
    Author            = 'Jason Ulbright'
    Description       = 'Vendor Version Monitor - compares MECM-packaged versions against vendor releases, with NVD CVE lookup.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        'Read-MonitorConfig'
        'Get-PackagerScripts'
        'Get-MecmApplicationVersions'
        'Invoke-VendorVersionCheck'
        'Compare-Versions'
        'Invoke-NvdCveQuery'
        'Invoke-NvdBatchQuery'
        'Read-NvdCache'
        'Write-NvdCache'
        'Export-VersionMonitorHtml'
        'Send-ReportNotification'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
