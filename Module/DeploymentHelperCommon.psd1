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

        # Search
        'Search-CMApplicationByName'
        'Search-CMCollectionByName'

        # DP Groups
        'Get-DPGroupList'
        'Start-ContentDistributionToGroups'
        'Invoke-ContentDistributionToGroups'
        'Get-ContentTargetedDPGroups'

        # Validation
        'Test-ApplicationExists'
        'Test-ContentDistributed'
        'Test-CollectionValid'
        'Test-CollectionSafe'
        'Test-DuplicateDeployment'
        'Get-DeploymentPreview'

        # SUG Validation
        'Test-SUGExists'

        # Execution
        'Invoke-ApplicationDeployment'
        'Invoke-SUGDeployment'
        'Invoke-PackageDeployment'
        'Invoke-TaskSequenceDeployment'

        # Packages
        'Search-CMPackageByName'
        'Test-PackageExists'
        'Get-CMPackagePrograms'
        'Test-DuplicatePackageDeployment'

        # Task Sequences
        'Search-CMTaskSequenceByName'
        'Test-TaskSequenceExists'
        'Test-DuplicateTaskSequenceDeployment'

        # Software Update Groups (extended)
        'Search-CMSoftwareUpdateGroupByName'
        'Test-DuplicateSUGDeployment'

        # Templates
        'Get-DeploymentTemplates'
        'Save-DeploymentTemplate'
        'Remove-DeploymentTemplate'

        # Deployment Log
        'Write-DeploymentLog'
        'Get-DeploymentHistory'

        # Export
        'Export-DeploymentHistoryCsv'
        'Export-DeploymentHistoryHtml'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
