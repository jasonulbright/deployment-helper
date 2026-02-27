@{
    RootModule        = 'DeploymentHelperCommon.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'c3d4e5f6-a7b8-9012-cdef-345678901234'
    Author            = 'Jason Ulbright'
    Description       = 'MECM application deployment with pre-execution validation, safety guardrails, and immutable audit logging.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging
        'Initialize-Logging'
        'Write-Log'

        # CM Connection
        'Connect-CMSite'
        'Disconnect-CMSite'
        'Test-CMConnection'

        # Validation
        'Test-ApplicationExists'
        'Test-ContentDistributed'
        'Test-CollectionValid'
        'Test-CollectionSafe'
        'Test-DuplicateDeployment'
        'Get-DeploymentPreview'

        # Execution
        'Invoke-ApplicationDeployment'

        # Deployment Log
        'Write-DeploymentLog'
        'Get-DeploymentHistory'

        # Templates
        'Get-DeploymentTemplates'

        # Export
        'Export-DeploymentHistoryCsv'
        'Export-DeploymentHistoryHtml'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
