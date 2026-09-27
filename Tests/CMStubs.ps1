<#
.SYNOPSIS
    Global stubs for the ConfigurationManager cmdlets that tests mock with a
    -ParameterFilter.

.DESCRIPTION
    A stub always replaces the command, also on a machine with the console
    installed: a mock built from the real cmdlet inherits its typed
    parameters (-Schedule and -InputObject take IResultObject), so a
    PSCustomObject test double fails to bind. A global function wins over a
    cmdlet of the same name. Parameter names match the cmdlet parameters the
    module passes. Value parameters stay untyped; a name in the switch list
    becomes [switch], because an untyped parameter demands an argument and
    a bare -Fast would fail to bind.
#>

$script:CMStubParameters = @{
    'Get-CMApplication'              = @('Name', 'Fast', 'DisableWildcardHandling')
    'Get-CMPackage'                  = @('Name', 'Fast', 'DisableWildcardHandling')
    'Get-CMTaskSequence'             = @('Name', 'Fast', 'DisableWildcardHandling')
    'Get-CMSoftwareUpdateGroup'      = @('Name', 'DisableWildcardHandling')
    'Get-CMCollection'               = @('Id', 'Name', 'CollectionType', 'DisableWildcardHandling')
    'Get-CMDeployment'               = @('DeploymentId', 'CollectionName', 'FeatureType')
    'Get-CMApplicationDeployment'    = @('Name', 'DeploymentId', 'CollectionName', 'Summary', 'DisableWildcardHandling')
    'Get-CMPackageDeployment'        = @('DeploymentId', 'PackageId', 'ProgramName', 'CollectionName', 'Summary', 'DisableWildcardHandling')
    'Get-CMTaskSequenceDeployment'   = @('DeploymentId', 'TaskSequenceId', 'CollectionName', 'Fast', 'Summary', 'DisableWildcardHandling')
    'Get-CMUpdateGroupDeployment'    = @('DeploymentId', 'Name', 'CollectionName', 'Summary', 'DisableWildcardHandling')
    'New-CMSchedule'                 = @('Start', 'Nonrecurring', 'IsUtc')
    'New-CMApplicationDeployment'    = @('Name', 'CollectionName', 'DeployPurpose', 'DeployAction', 'AvailableDateTime',
                                         'DeadlineDateTime', 'TimeBaseOn', 'UserNotification', 'OverrideServiceWindow',
                                         'RebootOutsideServiceWindow', 'UseMeteredNetwork')
    'New-CMSoftwareUpdateDeployment' = @('SoftwareUpdateGroupName', 'CollectionName', 'DeploymentType', 'AvailableDateTime',
                                         'DeadlineDateTime', 'TimeBasedOn', 'UserNotification', 'SoftwareInstallation',
                                         'AllowRestart', 'RequirePostRebootFullScan', 'ProtectedType', 'UnprotectedType',
                                         'UseMeteredNetwork', 'DownloadFromMicrosoftUpdate')
    'New-CMPackageDeployment'        = @('StandardProgram', 'PackageId', 'ProgramName', 'CollectionId', 'DeployPurpose',
                                         'AvailableDateTime', 'DeadlineDateTime', 'Schedule', 'FastNetworkOption',
                                         'SlowNetworkOption', 'RerunBehavior', 'UseUtcForAvailableSchedule',
                                         'UseUtcForExpireSchedule', 'SoftwareInstallation', 'SystemRestart', 'UseMeteredNetwork')
    'New-CMTaskSequenceDeployment'   = @('InputObject', 'CollectionId', 'DeployPurpose', 'AvailableDateTime', 'DeadlineDateTime',
                                         'Schedule', 'Availability', 'ShowTaskSequenceProgress', 'UseUtcForAvailableSchedule',
                                         'UseUtcForExpireSchedule', 'SoftwareInstallation', 'SystemRestart', 'UseMeteredNetwork')
}

$script:CMStubSwitches = @('Summary', 'Fast', 'Nonrecurring', 'IsUtc', 'StandardProgram', 'DisableWildcardHandling')

function Set-CMTestStub {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Defines test stub functions in the test session.')]
    param([Parameter(Mandatory)][string[]]$Name)
    foreach ($n in $Name) {
        $params = ($script:CMStubParameters[$n] | ForEach-Object {
            if ($_ -in $script:CMStubSwitches) { '[switch]$' + $_ } else { '$' + $_ }
        }) -join ', '
        Set-Item -Path "function:global:$n" -Value ([scriptblock]::Create("[CmdletBinding()] param($params)"))
    }
}
