@{
    RootModule        = 'DeploymentHelperCommon.psm1'
    ModuleVersion     = '2026.09.26.0010'
    GUID              = 'c3d4e5f6-a7b8-9012-cdef-345678901234'
    Author            = 'Jason Ulbright'
    Description       = 'Configuration Manager application deployment with pre-execution validation, safety guardrails, and immutable audit logging.'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Logging and CM connection come from the vendored SuiteCommon
        # module (Lib\SuiteCommon), imported globally by the root module.

        # Search
        'Search-CMApplicationByName'
        'Search-CMCollectionByName'

        # Browse
        'Get-CMBrowseList'
        'Get-CMCollectionFolderInfo'
        'Add-CollectionFolderId'
        'Select-BrowseMatch'

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
        'Test-CollectionIdBuiltIn'
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
        'Get-DuplicateCheckFailure'

        # Templates
        'Get-DeploymentTemplates'
        'Save-DeploymentTemplate'
        'Remove-DeploymentTemplate'

        # Deployment Log
        'Test-DeploymentLogWritable'
        'Write-DeploymentLog'
        'Get-DeploymentHistory'

        # Export
        'Export-DeploymentHistoryCsv'
        'Export-DeploymentHistoryHtml'

        # Ring plans
        'ConvertFrom-RingDateText'
        'ConvertTo-RingDateText'
        'Get-RingPlanSeed'
        'Initialize-RingPlanFolder'
        'ConvertTo-RingPlan'
        'Test-RingPlan'
        'Import-RingPlan'
        'Get-RingPlanList'
        'Get-RingLabel'
        'Expand-RingPlan'
        'Get-RingNow'
        'Test-RingExpansion'

        # Ring runs
        'New-RingRunId'
        'Get-RingRunFileName'
        'Get-RingObjectIdentity'
        'New-RingRun'
        'New-RingRunFile'
        'Save-RingRun'
        'Read-RingRun'
        'Get-RingRunRing'
        'Set-RingRunRingStatus'
        'Get-RingRunNextHeld'
        'Test-RingRunFinished'
        'Test-RingRunPromotable'
        'Get-RingPromoteShift'
        'Move-RingRunSchedule'
        'Enter-RingRunLock'
        'Exit-RingRunLock'
        'Get-RingDeploymentSummary'
        'Get-RingLiveSummary'
        'Update-RingRunReconcile'
        'Test-RingThreshold'
        'Get-RingRunFile'
        'Close-RingRun'
        'Test-RingPreflight'
        'Invoke-RingDeployment'
        'New-RingAuditRecord'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
