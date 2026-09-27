<#
.SYNOPSIS
    Core module for Deployment Helper.

.DESCRIPTION
    Import this module to get:
      - Structured logging and CM site connection management via the
        vendored SuiteCommon module (Lib\SuiteCommon)
      - Pre-execution validation (Test-ApplicationExists, Test-ContentDistributed, Test-CollectionValid, Test-CollectionSafe, Test-DuplicateDeployment)
      - DP group management (Get-DPGroupList, Start-ContentDistributionToGroups)
      - Deployment preview and execution (Get-DeploymentPreview, Invoke-ApplicationDeployment)
      - Immutable deployment audit log (Write-DeploymentLog, Get-DeploymentHistory)
      - Deployment templates (Get-DeploymentTemplates)
      - Export to CSV and HTML (Export-DeploymentHistoryCsv, Export-DeploymentHistoryHtml)
      - Ring plans and ring runs (Import-RingPlan, Expand-RingPlan, Test-RingExpansion,
        New-RingRun, Test-RingRunPromotable, Update-RingRunReconcile, Invoke-RingDeployment)

.EXAMPLE
    Import-Module "$PSScriptRoot\Module\DeploymentHelperCommon.psd1" -Force
    Initialize-Logging -LogPath "C:\temp\dh.log"
    Connect-CMSite -SiteCode 'MCM' -SMSProvider 'sccm.domain.com'
#>

# ---------------------------------------------------------------------------
# Shared core (vendored SuiteCommon)
# ---------------------------------------------------------------------------
# Logging (Initialize-Logging, Write-Log), CM connection (Connect-CMSite,
# Disconnect-CMSite, Test-CMConnection, Get-CMConnectionInfo), and settings
# persistence come from the vendored copy at Lib\SuiteCommon\. -Global makes
# the functions resolvable from the shell script and from this module alike;
# the guard keeps a -Force reimport of this module from resetting SuiteCommon
# state mid-session.
if (-not (Get-Module SuiteCommon)) {
    Import-Module (Join-Path $PSScriptRoot '..\Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking
}

# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

function Test-ApplicationExists {
    param([Parameter(Mandatory)][string]$ApplicationName)

    try {
        $app = Get-CMApplication -Name $ApplicationName -DisableWildcardHandling -ErrorAction Stop
        if ($null -eq $app) {
            Write-Log "Application not found: $ApplicationName" -Level WARN
            return $null
        }
        Write-Log "Application found: $ApplicationName v$($app.SoftwareVersion) (PackageID: $($app.PackageID))"
        return $app
    }
    catch {
        Write-Log "Error querying application '$ApplicationName': $_" -Level ERROR
        return $null
    }
}

function Search-CMApplicationByName {
    param([Parameter(Mandatory)][string]$SearchText)

    try {
        # @(...) wrap so a single-hit result is still an array; otherwise
        # $apps.Count is $null on exactly-one-match and the log line
        # reads "  result(s)" with the number missing.
        $apps = @(Get-CMApplication -Name "*$SearchText*" -Fast -ErrorAction Stop |
            Select-Object LocalizedDisplayName, SoftwareVersion, PackageID, DateLastModified |
            Sort-Object LocalizedDisplayName)
        Write-Log "Application search '$SearchText': $($apps.Count) result(s)"
        return $apps
    }
    catch {
        Write-Log "Error searching applications: $_" -Level ERROR
        return @()
    }
}

function Search-CMCollectionByName {
    param([Parameter(Mandatory)][string]$SearchText)

    try {
        $cols = @(Get-CMCollection -Name "*$SearchText*" -CollectionType Device -ErrorAction Stop |
            Select-Object Name, CollectionID, MemberCount, LastRefreshTime |
            Sort-Object Name)
        Write-Log "Collection search '$SearchText': $($cols.Count) result(s)"
        return $cols
    }
    catch {
        Write-Log "Error searching collections: $_" -Level ERROR
        return @()
    }
}

# ---------------------------------------------------------------------------
# Browse (bulk load + client-side filter)
# ---------------------------------------------------------------------------

function Get-CMBrowseList {
    <#
    .SYNOPSIS
        Returns every object of one type as flat rows for the browse dialogs.
        One provider read per call; the dialogs filter the rows locally.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Apps', 'Packages', 'TaskSequences', 'SUG', 'Collections')]
        [string]$Type
    )

    $rows = switch ($Type) {
        'Apps' {
            @(Get-CMApplication -Fast -ErrorAction Stop |
                Select-Object LocalizedDisplayName, SoftwareVersion, PackageID, DateLastModified |
                Sort-Object LocalizedDisplayName)
        }
        'Packages' {
            @(Get-CMPackage -Fast -ErrorAction Stop |
                Select-Object Name, PackageID, Manufacturer, Version |
                Sort-Object Name)
        }
        'TaskSequences' {
            @(Get-CMTaskSequence -Fast -ErrorAction Stop |
                Select-Object Name, PackageID, BootImageID, Description |
                Sort-Object Name)
        }
        'SUG' {
            @(Get-CMSoftwareUpdateGroup -ErrorAction Stop |
                Select-Object LocalizedDisplayName, NumberOfUpdates, NumberOfExpiredUpdates, DateCreated |
                Sort-Object LocalizedDisplayName)
        }
        'Collections' {
            @(Get-CMCollection -CollectionType Device -ErrorAction Stop | ForEach-Object {
                $builtIn = if ($null -ne $_.IsBuiltIn) { [bool]$_.IsBuiltIn } else { ([string]$_.CollectionID) -like 'SMS*' }
                [PSCustomObject]@{
                    Name            = [string]$_.Name
                    CollectionID    = [string]$_.CollectionID
                    MemberCount     = [int]$_.MemberCount
                    LastRefreshTime = $_.LastRefreshTime
                    IsBuiltIn       = $builtIn
                }
            } | Sort-Object Name)
        }
    }

    $rows = @($rows)
    Write-Log "Browse load '$Type': $($rows.Count) object(s)"
    return ,$rows
}

function Get-CMCollectionFolderInfo {
    <#
    .SYNOPSIS
        Returns the device-collection console folders and the
        CollectionID -> folder map. Two CIM reads, independent of row count.
    #>
    param(
        [Parameter(Mandatory)][string]$SMSProvider,
        [Parameter(Mandatory)][string]$SiteCode
    )

    $namespace = "root\sms\site_$SiteCode"

    $folders = @(
        Get-CimInstance -ComputerName $SMSProvider -Namespace $namespace `
            -ClassName SMS_ObjectContainerNode -Filter 'ObjectType = 5000' -ErrorAction Stop |
        ForEach-Object {
            [PSCustomObject]@{
                FolderID = [int]$_.ContainerNodeID
                Name     = [string]$_.Name
                ParentID = [int]$_.ParentContainerNodeID
            }
        }
    )

    $items = @(
        Get-CimInstance -ComputerName $SMSProvider -Namespace $namespace `
            -ClassName SMS_ObjectContainerItem -Filter 'ObjectType = 5000' -ErrorAction Stop
    )
    $map = @{}
    foreach ($i in $items) { $map[[string]$i.InstanceKey] = [int]$i.ContainerNodeID }

    Write-Log "Collection folders: $($folders.Count) folder(s), $($map.Count) placed collection(s)"
    return @{ Folders = $folders; FolderMap = $map }
}

function Add-CollectionFolderId {
    <#
    .SYNOPSIS
        Adds a FolderID property to each collection row. A collection absent
        from the map lives at the tree root (FolderID 0).
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Collections,
        [hashtable]$FolderMap = @{}
    )

    $out = @(foreach ($c in $Collections) {
        $key = [string]$c.CollectionID
        $fid = if ($FolderMap.ContainsKey($key)) { [int]$FolderMap[$key] } else { 0 }
        $c | Add-Member -MemberType NoteProperty -Name FolderID -Value $fid -Force -PassThru
    })
    return ,$out
}

function Select-BrowseMatch {
    <#
    .SYNOPSIS
        Filters browse rows by a case-insensitive substring over the given
        properties (default: every property). An empty needle returns all
        rows in their original order.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Items,
        [AllowEmptyString()][AllowNull()][string]$Needle,
        [string[]]$Property
    )

    $needle = ([string]$Needle).Trim()
    if ($needle.Length -eq 0) { return ,@($Items) }

    $out = @(foreach ($item in $Items) {
        $names = if ($Property -and $Property.Count -gt 0) { $Property } else { @($item.PSObject.Properties.Name) }
        $hit = $false
        foreach ($n in $names) {
            $v = $item.$n
            if ($null -eq $v) { continue }
            if (([string]$v).IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $hit = $true; break }
        }
        if ($hit) { $item }
    })
    return ,$out
}

function Test-ContentDistributed {
    param([Parameter(Mandatory)]$Application)

    try {
        $status = Get-CMDistributionStatus -Id $Application.PackageID -ErrorAction Stop
        if ($null -eq $status) {
            Write-Log "No distribution status found for $($Application.LocalizedDisplayName) - content may not be distributed to any DP" -Level WARN
            return @{ Targeted = 0; NumberSuccess = 0; NumberInProgress = 0; NumberErrors = 0; IsFullyDistributed = $false }
        }

        $result = @{
            Targeted           = $status.Targeted
            NumberSuccess      = $status.NumberSuccess
            NumberInProgress   = $status.NumberInProgress
            NumberErrors       = $status.NumberErrors
            IsFullyDistributed = ($status.NumberSuccess -ge $status.Targeted -and $status.Targeted -gt 0 -and $status.NumberErrors -eq 0)
        }

        if ($result.IsFullyDistributed) {
            Write-Log "Content fully distributed: $($result.NumberSuccess)/$($result.Targeted) DPs"
        } else {
            Write-Log ("Content NOT fully distributed: {0}/{1} success, {2} errors, {3} in progress" -f
                $result.NumberSuccess, $result.Targeted, $result.NumberErrors, $result.NumberInProgress) -Level WARN
        }
        return $result
    }
    catch {
        Write-Log "Error checking distribution status: $_" -Level ERROR
        return @{ Targeted = 0; NumberSuccess = 0; NumberInProgress = 0; NumberErrors = 0; IsFullyDistributed = $false; Error = $_.ToString() }
    }
}

function Get-DPGroupList {
    try {
        $groups = Get-CMDistributionPointGroup -ErrorAction Stop | Sort-Object Name
        Write-Log "Retrieved $($groups.Count) DP group(s)"
        return $groups
    }
    catch {
        Write-Log "Error retrieving DP groups: $_" -Level ERROR
        return @()
    }
}

function Start-ContentDistributionToGroups {
    param(
        [Parameter(Mandatory)]$Application,
        [Parameter(Mandatory)][string[]]$DPGroupNames
    )

    $results = @()
    foreach ($groupName in $DPGroupNames) {
        try {
            Start-CMContentDistribution -ApplicationName $Application.LocalizedDisplayName -DistributionPointGroupName $groupName -ErrorAction Stop
            Write-Log "Content distribution started to DP group '$groupName'"
            $results += @{ Group = $groupName; Success = $true }
        }
        catch {
            if ($_.Exception.Message -match 'already been targeted') {
                Write-Log "Content already distributed to DP group '$groupName'" -Level INFO
                $results += @{ Group = $groupName; Success = $true; AlreadyTargeted = $true }
            } else {
                Write-Log "Error distributing to DP group '$groupName': $_" -Level ERROR
                $results += @{ Group = $groupName; Success = $false; Error = $_.ToString() }
            }
        }
    }
    return $results
}

function Invoke-ContentDistributionToGroups {
    <#
    .SYNOPSIS
        Distribute content (application / package / task sequence) to one or
        more DP groups. Polymorphic; dispatches by -Type.

    .DESCRIPTION
        Wraps Start-CMContentDistribution. 'already been targeted' is treated
        as an informational success so re-submitting a DP group is harmless.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('Application','Package','TaskSequence')][string]$Type,
        [Parameter(Mandatory)]$TargetObject,
        [Parameter(Mandatory)][string[]]$DPGroupNames
    )

    $results = @()
    foreach ($groupName in $DPGroupNames) {
        $p = @{ DistributionPointGroupName = $groupName; ErrorAction = 'Stop' }
        switch ($Type) {
            'Application'  { $p['ApplicationName']  = $TargetObject.LocalizedDisplayName }
            'Package'      { $p['PackageId']         = $TargetObject.PackageID }
            'TaskSequence' { $p['TaskSequenceId']    = $TargetObject.PackageID }
        }
        try {
            Start-CMContentDistribution @p
            Write-Log "Content distribution started to DP group '$groupName' ($Type)"
            $results += @{ Group = $groupName; Success = $true }
        }
        catch {
            $msg = $_.Exception.Message
            # Different CM builds phrase "already distributed" differently; accept a
            # broad match rather than brittle exact wording.
            if ($msg -match 'already been targeted' -or
                $msg -match 'already been distributed' -or
                $msg -match 'No content destination.*already been distributed') {
                Write-Log "DP group '$groupName' already has this content"
                $results += @{ Group = $groupName; Success = $true; AlreadyTargeted = $true }
            } else {
                Write-Log "Error distributing to DP group '$groupName': $_" -Level ERROR
                $results += @{ Group = $groupName; Success = $false; Error = $_.ToString() }
            }
        }
    }
    return $results
}

function Get-ContentTargetedDPGroups {
    <#
    .SYNOPSIS
        Returns the list of DP group names that currently hold this content.

    .DESCRIPTION
        Queries SMS_DPGroupContentInfo against the connected site's WMI
        namespace. ObjectID format differs by content type:
          - Packages, Task Sequences, OS Images, Boot Images, Driver Packages:
            PackageID (e.g., "MCM00289")
          - Applications: CI_UniqueID / ModelName (e.g.,
            "ScopeId_.../Application_...")
        Pass the right identifier for the target type. For a convenience
        helper that picks the right identifier automatically, see the
        calling code in start-deploymenthelper.ps1.

        Falls back to an empty list with a WARN on any error so the UI can
        still proceed with "nothing pre-checked". Requires an active
        Connect-CMSite; ConnectedSiteCode + ConnectedSMSProvider are set
        by Connect-CMSite in this module.
    #>
    param(
        [Parameter(Mandatory)][string]$ObjectID
    )

    if (-not $script:ConnectedSiteCode -or -not $script:ConnectedSMSProvider) {
        return @()
    }

    try {
        $ns = "root\SMS\site_$($script:ConnectedSiteCode)"

        # Skip -ComputerName when the provider is the local machine: avoids the
        # second-hop auth failure ("specified logon session does not exist") that
        # happens when WMI is asked to remote-authenticate back to itself from
        # inside a PSSession. Works transparently whether the GUI runs on the
        # site server itself or on a separate engineer box with AdminUI.
        $provider = [string]$script:ConnectedSMSProvider
        $isLocal = $false
        if ($provider) {
            $localNames = @($env:COMPUTERNAME, "$env:COMPUTERNAME.$env:USERDNSDOMAIN", 'localhost', '.')
            foreach ($n in $localNames) {
                if ($n -and $provider -ieq $n) { $isLocal = $true; break }
            }
        }
        $common = @{ Namespace = $ns; ErrorAction = 'Stop' }
        if (-not $isLocal) { $common['ComputerName'] = $provider }

        # WQL filter needs backslashes in CI_UniqueID app IDs escaped: they aren't
        # in the ModelName, but forward slashes are. Escape single quotes just in case.
        $esc = $ObjectID.Replace("'", "''")
        $entries = Get-CimInstance @common -ClassName SMS_DPGroupContentInfo -Filter "ObjectID='$esc'"

        $names = @()
        foreach ($e in $entries) {
            $common2 = @{ Namespace = $ns; ErrorAction = 'SilentlyContinue' }
            if (-not $isLocal) { $common2['ComputerName'] = $provider }
            $g = Get-CimInstance @common2 -ClassName SMS_DistributionPointGroup -Filter "GroupID='$($e.GroupID)'"
            if ($g) { $names += [string]$g.Name }
        }
        $unique = @($names | Sort-Object -Unique)
        Write-Log "Content '$ObjectID' is already targeted to DP groups: $($unique -join ', ')"
        return $unique
    }
    catch {
        Write-Log "Get-ContentTargetedDPGroups WMI query failed: $_" -Level WARN
        return @()
    }
}

function Test-CollectionValid {
    [CmdletBinding(DefaultParameterSetName = 'ByName')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'ByName')][string]$CollectionName,
        [Parameter(Mandatory, ParameterSetName = 'ById')][string]$CollectionId
    )

    $label = if ($PSCmdlet.ParameterSetName -eq 'ById') { $CollectionId } else { $CollectionName }
    try {
        $col = if ($PSCmdlet.ParameterSetName -eq 'ById') {
            Get-CMCollection -Id $CollectionId -ErrorAction Stop
        } else {
            Get-CMCollection -Name $CollectionName -DisableWildcardHandling -ErrorAction Stop
        }
        if ($null -eq $col) {
            Write-Log "Collection not found: $label" -Level WARN
            return $null
        }
        if ($col.CollectionType -ne 2) {
            Write-Log "Collection '$label' is a User collection, not Device. Deployment requires a Device collection." -Level WARN
            return $null
        }
        Write-Log "Collection found: $($col.Name) (ID: $($col.CollectionID), Members: $($col.MemberCount))"
        return $col
    }
    catch {
        Write-Log "Error querying collection '$label': $_" -Level ERROR
        return $null
    }
}

# Built-in collections carry the reserved "SMS" prefix (SMS00001 All Systems,
# SMSDM003 All Desktop and Server Clients). No site can use the site code
# SMS, so the prefix never matches a custom collection.
$script:BuiltInCollectionPrefixes = @('SMSDM', 'SMS')

function Test-CollectionIdBuiltIn {
    param([AllowEmptyString()][AllowNull()][string]$CollectionId)

    $id = ([string]$CollectionId).Trim()
    if ($id.Length -eq 0) { return $false }
    foreach ($p in $script:BuiltInCollectionPrefixes) {
        if ($id.StartsWith($p, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Test-CollectionSafe {
    param([Parameter(Mandatory)]$Collection)

    $collectionId = [string]$Collection.CollectionID

    if ([string]::IsNullOrWhiteSpace($collectionId)) {
        Write-Log "BLOCKED: Collection '$($Collection.Name)' has no CollectionID. Deployment not allowed." -Level ERROR
        return @{ IsSafe = $false; Reason = 'The collection has no CollectionID, so the built-in check cannot run.' }
    }
    if (Test-CollectionIdBuiltIn -CollectionId $collectionId) {
        Write-Log "BLOCKED: Collection '$($Collection.Name)' ($collectionId) is a built-in system collection. Deployment not allowed." -Level ERROR
        return @{ IsSafe = $false; Reason = "Built-in system collection ($collectionId) is blocked for safety." }
    }

    Write-Log "Collection '$($Collection.Name)' ($collectionId) passed safety check"
    return @{ IsSafe = $true; Reason = '' }
}

function Get-UnsafeCollectionResult {
    <#
    .SYNOPSIS
        Returns a failed deployment result when the target is not safe, or
        $null when it is. Every Invoke-*Deployment function calls this before
        its New-CM* cmdlet.
    #>
    param([Parameter(Mandatory)]$Collection)

    $safe = Test-CollectionSafe -Collection $Collection
    if ($safe.IsSafe) { return $null }
    return @{
        Success            = $false
        DeploymentID       = $null
        DeploymentUniqueID = $null
        Error              = $safe.Reason
    }
}

function Test-DuplicateDeployment {
    param(
        [Parameter(Mandatory)][string]$ApplicationName,
        [Parameter(Mandatory)][string]$CollectionName
    )

    try {
        $existing = Get-CMApplicationDeployment -Name $ApplicationName -CollectionName $CollectionName -DisableWildcardHandling -ErrorAction Stop
        if ($null -ne $existing -and @($existing).Count -gt 0) {
            Write-Log "Duplicate deployment found: '$ApplicationName' already deployed to '$CollectionName'" -Level WARN
            return $existing
        }
        Write-Log "No duplicate deployment: '$ApplicationName' to '$CollectionName'"
        return $null
    }
    catch {
        Write-Log "Error checking for duplicate deployment: $_" -Level ERROR
        return [PSCustomObject]@{ DuplicateCheckFailed = $true; Error = $_.Exception.Message }
    }
}

function Get-DeploymentPreview {
    param(
        [Parameter(Mandatory)]$TargetObject,
        [Parameter(Mandatory)]$Collection,
        [string]$DeploymentType = 'Application'
    )

    if ($DeploymentType -eq 'SUG') {
        return @{
            ApplicationName    = $TargetObject.LocalizedDisplayName
            ApplicationVersion = "($($TargetObject.NumberOfUpdates) updates)"
            CollectionName     = $Collection.Name
            CollectionID       = $Collection.CollectionID
            MemberCount        = $Collection.MemberCount
        }
    } else {
        return @{
            ApplicationName    = $TargetObject.LocalizedDisplayName
            ApplicationVersion = $TargetObject.SoftwareVersion
            CollectionName     = $Collection.Name
            CollectionID       = $Collection.CollectionID
            MemberCount        = $Collection.MemberCount
        }
    }
}

# ---------------------------------------------------------------------------
# SUG Validation
# ---------------------------------------------------------------------------

function Test-SUGExists {
    param([Parameter(Mandatory)][string]$SUGName)

    try {
        $sug = Get-CMSoftwareUpdateGroup -Name $SUGName -DisableWildcardHandling -ErrorAction Stop
        if ($null -eq $sug) {
            Write-Log "Software Update Group not found: $SUGName" -Level WARN
            return $null
        }
        Write-Log "SUG found: $SUGName ($($sug.NumberOfUpdates) updates, $($sug.NumberOfExpiredUpdates) expired)"
        if ($sug.NumberOfUpdates -eq 0) {
            Write-Log "SUG '$SUGName' contains 0 updates" -Level WARN
        }
        return $sug
    }
    catch {
        Write-Log "Error querying SUG '$SUGName': $_" -Level ERROR
        return $null
    }
}

# ---------------------------------------------------------------------------
# Execution
# ---------------------------------------------------------------------------

function Invoke-ApplicationDeployment {
    param(
        [Parameter(Mandatory)]$Application,
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)][ValidateSet('Required','Available')][string]$DeployPurpose,
        [Parameter(Mandatory)][datetime]$AvailableDateTime,
        [datetime]$DeadlineDateTime,
        [ValidateSet('LocalTime','Utc')][string]$TimeBasedOn = 'LocalTime',
        [ValidateSet('DisplayAll','DisplaySoftwareCenterOnly','HideAll')]
        [string]$UserNotification = 'DisplayAll',
        [bool]$OverrideServiceWindow = $false,
        [bool]$RebootOutsideServiceWindow = $false,
        [bool]$UseMeteredNetwork = $false
    )

    $blocked = Get-UnsafeCollectionResult -Collection $Collection
    if ($blocked) { return $blocked }

    $params = @{
        Name              = $Application.LocalizedDisplayName
        CollectionName    = $Collection.Name
        DeployPurpose     = $DeployPurpose
        DeployAction      = 'Install'
        AvailableDateTime = $AvailableDateTime
        TimeBaseOn        = $TimeBasedOn
        UserNotification  = $UserNotification
        ErrorAction       = 'Stop'
    }

    if ($DeployPurpose -eq 'Required') {
        if ($DeadlineDateTime) { $params['DeadlineDateTime'] = $DeadlineDateTime }
        $params['OverrideServiceWindow']      = $OverrideServiceWindow
        $params['RebootOutsideServiceWindow'] = $RebootOutsideServiceWindow
    }
    if ($UseMeteredNetwork) {
        $params['UseMeteredNetwork'] = $true
    }

    try {
        Write-Log ("Executing deployment: {0} v{1} -> {2} ({3} devices) as {4}" -f
            $Application.LocalizedDisplayName, $Application.SoftwareVersion,
            $Collection.Name, $Collection.MemberCount, $DeployPurpose)

        $deployment = New-CMApplicationDeployment @params

        Write-Log "Deployment created successfully (ID: $($deployment.AssignmentID))"
        return @{
            Success            = $true
            DeploymentID       = $deployment.AssignmentID
            DeploymentUniqueID = [string]$deployment.AssignmentUniqueID
            Error              = $null
        }
    }
    catch {
        Write-Log "Deployment FAILED: $_" -Level ERROR
        return @{
            Success            = $false
            DeploymentID       = $null
            DeploymentUniqueID = $null
            Error              = $_.ToString()
        }
    }
}

function Invoke-SUGDeployment {
    param(
        [Parameter(Mandatory)]$SUG,
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)][ValidateSet('Required','Available')][string]$DeployPurpose,
        [Parameter(Mandatory)][datetime]$AvailableDateTime,
        [datetime]$DeadlineDateTime,
        [ValidateSet('LocalTime','Utc')][string]$TimeBasedOn = 'LocalTime',
        [ValidateSet('DisplayAll','DisplaySoftwareCenterOnly','HideAll')]
        [string]$UserNotification = 'DisplayAll',
        [bool]$SoftwareInstallation = $false,
        [bool]$AllowRestart = $false,
        [bool]$UseMeteredNetwork = $false,
        [bool]$AllowBoundaryFallback = $true,
        [bool]$DownloadFromMicrosoftUpdate = $false,
        [bool]$RequirePostRebootFullScan = $true
    )

    $blocked = Get-UnsafeCollectionResult -Collection $Collection
    if ($blocked) { return $blocked }

    $params = @{
        SoftwareUpdateGroupName    = $SUG.LocalizedDisplayName
        CollectionName             = $Collection.Name
        DeploymentType             = $DeployPurpose
        AvailableDateTime          = $AvailableDateTime
        TimeBasedOn                = $TimeBasedOn
        UserNotification           = $UserNotification
        SoftwareInstallation       = $SoftwareInstallation
        AllowRestart               = $AllowRestart
        RequirePostRebootFullScan  = $RequirePostRebootFullScan
        ErrorAction                = 'Stop'
    }

    if ($DeployPurpose -eq 'Required' -and $DeadlineDateTime) {
        $params['DeadlineDateTime'] = $DeadlineDateTime
    }

    # Required SUG: download fallback settings
    if ($DeployPurpose -eq 'Required') {
        $params['ProtectedType']   = 'RemoteDistributionPoint'
        $params['UnprotectedType'] = if ($AllowBoundaryFallback) { 'UnprotectedDistributionPoint' } else { 'NoInstall' }
    }

    if ($UseMeteredNetwork) {
        $params['UseMeteredNetwork'] = $true
    }
    if ($DownloadFromMicrosoftUpdate) {
        $params['DownloadFromMicrosoftUpdate'] = $true
    }

    try {
        Write-Log ("Executing SUG deployment: {0} ({1} updates) -> {2} ({3} devices) as {4}" -f
            $SUG.LocalizedDisplayName, $SUG.NumberOfUpdates,
            $Collection.Name, $Collection.MemberCount, $DeployPurpose)

        $deployment = New-CMSoftwareUpdateDeployment @params

        Write-Log "SUG deployment created successfully (ID: $($deployment.AssignmentID))"
        return @{
            Success            = $true
            DeploymentID       = $deployment.AssignmentID
            DeploymentUniqueID = [string]$deployment.AssignmentUniqueID
            Error              = $null
        }
    }
    catch {
        Write-Log "SUG deployment FAILED: $_" -Level ERROR
        return @{
            Success            = $false
            DeploymentID       = $null
            DeploymentUniqueID = $null
            Error              = $_.ToString()
        }
    }
}

function Save-DeploymentTemplate {
    param(
        [Parameter(Mandatory)][string]$TemplatePath,
        [Parameter(Mandatory)][string]$TemplateName,
        [Parameter(Mandatory)][hashtable]$Config
    )

    $parentDir = Split-Path -Path $TemplatePath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $template = [ordered]@{
        Name                        = $TemplateName
        TargetCollectionName        = [string]$Config.TargetCollectionName
        TargetCollectionID          = [string]$Config.TargetCollectionID
        DeployPurpose               = $Config.DeployPurpose
        UserNotification            = $Config.UserNotification
        TimeBasedOn                 = if ($Config.TimeBasedOn) { $Config.TimeBasedOn } else { 'LocalTime' }
        OverrideServiceWindow       = $Config.OverrideServiceWindow
        RebootOutsideServiceWindow  = $Config.RebootOutsideServiceWindow
        AllowMeteredConnection      = $Config.AllowMeteredConnection
        AllowBoundaryFallback       = if ($null -ne $Config.AllowBoundaryFallback) { $Config.AllowBoundaryFallback } else { $true }
        AllowMicrosoftUpdate        = if ($null -ne $Config.AllowMicrosoftUpdate) { $Config.AllowMicrosoftUpdate } else { $false }
        RequirePostRebootFullScan   = if ($null -ne $Config.RequirePostRebootFullScan) { $Config.RequirePostRebootFullScan } else { $true }
        DefaultDeadlineOffsetHours  = $Config.DefaultDeadlineOffsetHours
    }

    $template | ConvertTo-Json | Set-Content -LiteralPath $TemplatePath -Encoding UTF8
    Write-Log "Saved deployment template '$TemplateName' to $TemplatePath"
}

function Remove-DeploymentTemplate {
    param(
        [Parameter(Mandatory)][string]$TemplatePath
    )

    if (-not (Test-Path -LiteralPath $TemplatePath)) {
        Write-Log "Template to remove not found: $TemplatePath" -Level WARN
        return
    }
    Remove-Item -LiteralPath $TemplatePath -Force
    Write-Log "Removed deployment template: $TemplatePath"
}

# ---------------------------------------------------------------------------
# Deployment Log (JSONL - one JSON object per line, append-only)
# ---------------------------------------------------------------------------

function Test-DeploymentLogWritable {
    param([Parameter(Mandatory)][string]$LogPath)

    $stream = $null
    try {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force -ErrorAction Stop | Out-Null
        }
        # FileShare.None matches Add-Content on PS 5.1, which fails while any
        # reader holds the log open; a looser probe passes and the append
        # after the create then fails.
        $stream = [System.IO.File]::Open($LogPath, [System.IO.FileMode]::OpenOrCreate,
            [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        $stream.Flush()
        $stream.Dispose()
        $stream = $null
        return @{ Success = $true; Error = $null }
    }
    catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    }
    finally {
        if ($null -ne $stream) { try { $stream.Dispose() } catch { $null = $_ } }
    }
}

function Write-DeploymentLog {
    param(
        [Parameter(Mandatory)][string]$LogPath,
        [Parameter(Mandatory)][hashtable]$Record
    )

    try {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force -ErrorAction Stop | Out-Null
        }

        $entry = [ordered]@{
            Timestamp          = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
            User               = "$env:USERDOMAIN\$env:USERNAME"
            DeploymentType     = $Record.DeploymentType
            ApplicationName    = $Record.ApplicationName
            ApplicationVersion = $Record.ApplicationVersion
            CollectionName     = $Record.CollectionName
            CollectionID       = $Record.CollectionID
            MemberCount        = $Record.MemberCount
            DeployPurpose      = $Record.DeployPurpose
            DeployAction       = 'Install'
            DeadlineDateTime   = $Record.DeadlineDateTime
            DeploymentID       = $Record.DeploymentID
            Result             = $Record.Result
        }
        if ($Record.ContainsKey('PlanName')) {
            $entry['PlanName']  = $Record.PlanName
            $entry['RingIndex'] = $Record.RingIndex
            $entry['RingName']  = $Record.RingName
            $entry['RunId']     = $Record.RunId
        }

        $json = $entry | ConvertTo-Json -Compress
        Add-Content -LiteralPath $LogPath -Value $json -Encoding UTF8 -ErrorAction Stop
        Write-Log "Deployment log entry written to $LogPath"
        return @{ Success = $true; Error = $null }
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Log ("Audit log write failed ({0}): {1}" -f $LogPath, $errorMessage) -Level WARN
        return @{ Success = $false; Error = $errorMessage }
    }
}

function Get-DeploymentHistory {
    param([Parameter(Mandatory)][string]$LogPath)

    if (-not (Test-Path -LiteralPath $LogPath)) {
        Write-Log "Deployment log not found at $LogPath" -Level WARN
        return @()
    }

    $records = @()
    $lines = Get-Content -LiteralPath $LogPath -Encoding UTF8
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try {
            $records += ($line | ConvertFrom-Json)
        } catch {
            Write-Log "Skipped malformed log entry" -Level WARN
        }
    }

    Write-Log "Loaded $($records.Count) deployment history records"
    return $records
}

# ---------------------------------------------------------------------------
# Templates
# ---------------------------------------------------------------------------

function Get-DeploymentTemplates {
    param([Parameter(Mandatory)][string]$TemplatePath)

    if (-not (Test-Path -LiteralPath $TemplatePath)) {
        Write-Log "Templates folder not found: $TemplatePath" -Level WARN
        return @()
    }

    $templates = @()
    $files = Get-ChildItem -LiteralPath $TemplatePath -Filter '*.json' -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        try {
            $t = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json
            $templates += $t
        } catch {
            Write-Log "Failed to parse template $($f.Name): $_" -Level WARN
        }
    }

    Write-Log "Loaded $($templates.Count) deployment templates"
    return $templates
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

function Export-DeploymentHistoryCsv {
    param(
        [Parameter(Mandatory)][array]$Records,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    # Export-Csv takes its columns from the first record only; ring records
    # carry four extra fields, so the union keeps them in a mixed log.
    $columns = New-Object System.Collections.Generic.List[string]
    foreach ($r in $Records) {
        foreach ($n in $r.PSObject.Properties.Name) {
            if (-not $columns.Contains($n)) { $columns.Add($n) }
        }
    }
    $Records | Select-Object -Property @($columns) | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Exported deployment history CSV to $OutputPath"
}

function Export-DeploymentHistoryHtml {
    param(
        [Parameter(Mandatory)][array]$Records,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $css = @(
        '<style>',
        '  body { font-family: "Segoe UI", sans-serif; margin: 20px; background: #f8f9fa; }',
        '  h1 { color: #0078D4; margin-bottom: 4px; }',
        '  .subtitle { color: #666; margin-bottom: 16px; }',
        '  table { border-collapse: collapse; width: 100%; background: white; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }',
        '  th { background: #0078D4; color: white; padding: 10px 12px; text-align: left; font-size: 13px; }',
        '  td { padding: 8px 12px; border-bottom: 1px solid #e0e0e0; font-size: 13px; }',
        '  tr:nth-child(even) { background: #f5f7fa; }',
        '  tr:hover { background: #e8f0fe; }',
        '  .success { font-weight: bold; }',
        '  .failed { font-weight: bold; }',
        '  .success::before { content: "\2713  "; }',
        '  .failed::before  { content: "\2717  "; }',
        '</style>'
    ) -join "`r`n"

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $headerHtml = "<h1>Deployment History Report</h1><div class='subtitle'>Generated: $timestamp</div>"

    $columns = @('Timestamp','User','DeploymentType','ApplicationName','ApplicationVersion','CollectionName','MemberCount','DeployPurpose','DeadlineDateTime','DeploymentID','Result','PlanName','RingName')
    $thRow = ($columns | ForEach-Object { "<th>$_</th>" }) -join ""

    $bodyRows = foreach ($rec in $Records) {
        $cells = foreach ($col in $columns) {
            $val = $rec.$col
            if ($col -eq 'Result') {
                $cls = if ($val -match '^Success') { 'success' } else { 'failed' }
                "<td class='$cls'>$val</td>"
            } else {
                "<td>$val</td>"
            }
        }
        "<tr>$($cells -join '')</tr>"
    }

    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="UTF-8"><title>Deployment History Report</title>',
        $css,
        '</head><body>',
        $headerHtml,
        '<table>',
        "<tr>$thRow</tr>",
        ($bodyRows -join "`r`n"),
        '</table>',
        '</body></html>'
    ) -join "`r`n"

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    Write-Log "Exported deployment history HTML to $OutputPath"
}

# ---------------------------------------------------------------------------
# Packages (legacy Package + Program deployment)
# ---------------------------------------------------------------------------

function Search-CMPackageByName {
    param([Parameter(Mandatory)][string]$SearchText)

    try {
        $pkgs = @(Get-CMPackage -Name "*$SearchText*" -Fast -ErrorAction Stop |
            Select-Object Name, PackageID, Manufacturer, Version |
            Sort-Object Name)
        Write-Log "Package search '$SearchText': $($pkgs.Count) result(s)"
        return $pkgs
    }
    catch {
        Write-Log "Error searching packages: $_" -Level ERROR
        return @()
    }
}

function Test-PackageExists {
    param([Parameter(Mandatory)][string]$PackageName)

    try {
        $pkg = Get-CMPackage -Name $PackageName -Fast -DisableWildcardHandling -ErrorAction Stop
        if ($null -eq $pkg) {
            Write-Log "Package not found: $PackageName" -Level WARN
            return $null
        }
        Write-Log "Package found: $PackageName (PackageID: $($pkg.PackageID))"
        return $pkg
    }
    catch {
        Write-Log "Error querying package '$PackageName': $_" -Level ERROR
        return $null
    }
}

function Get-CMPackagePrograms {
    param([Parameter(Mandatory)]$Package)

    try {
        $programs = Get-CMProgram -PackageId $Package.PackageID -ErrorAction Stop |
            Select-Object ProgramName, CommandLine, PackageID |
            Sort-Object ProgramName
        Write-Log "Programs for package $($Package.PackageID): $($programs.Count)"
        return $programs
    }
    catch {
        Write-Log "Error listing programs for package '$($Package.Name)': $_" -Level ERROR
        return @()
    }
}

function Test-DuplicatePackageDeployment {
    param(
        [Parameter(Mandatory)][string]$PackageID,
        [Parameter(Mandatory)][string]$ProgramName,
        [Parameter(Mandatory)][string]$CollectionName
    )

    try {
        $existing = Get-CMPackageDeployment -PackageId $PackageID -ProgramName $ProgramName -CollectionName $CollectionName -DisableWildcardHandling -ErrorAction Stop
        if ($null -ne $existing -and @($existing).Count -gt 0) {
            Write-Log "Duplicate package deployment: PackageID=$PackageID Program='$ProgramName' Collection='$CollectionName'" -Level WARN
            return $existing
        }
        Write-Log "No duplicate package deployment: PackageID=$PackageID Program='$ProgramName' Collection='$CollectionName'"
        return $null
    }
    catch {
        Write-Log "Error checking duplicate package deployment: $_" -Level ERROR
        return [PSCustomObject]@{ DuplicateCheckFailed = $true; Error = $_.Exception.Message }
    }
}

# ---------------------------------------------------------------------------
# Software Update Groups (search + duplicate check)
# ---------------------------------------------------------------------------

function Search-CMSoftwareUpdateGroupByName {
    param([Parameter(Mandatory)][string]$SearchText)

    try {
        $sugs = @(Get-CMSoftwareUpdateGroup -Name "*$SearchText*" -ErrorAction Stop |
            Select-Object LocalizedDisplayName, NumberOfUpdates, NumberOfExpiredUpdates, DateCreated |
            Sort-Object LocalizedDisplayName)
        Write-Log "SUG search '$SearchText': $($sugs.Count) result(s)"
        return $sugs
    }
    catch {
        Write-Log "Error searching SUGs: $_" -Level ERROR
        return @()
    }
}

function Test-DuplicateSUGDeployment {
    param(
        [Parameter(Mandatory)][string]$SUGName,
        [Parameter(Mandatory)][string]$CollectionName
    )

    try {
        $existing = Get-CMUpdateGroupDeployment -Name $SUGName -CollectionName $CollectionName -DisableWildcardHandling -ErrorAction Stop
        if ($null -ne $existing -and @($existing).Count -gt 0) {
            Write-Log "Duplicate SUG deployment: SUG='$SUGName' Collection='$CollectionName'" -Level WARN
            return $existing
        }
        Write-Log "No duplicate SUG deployment: SUG='$SUGName' Collection='$CollectionName'"
        return $null
    }
    catch {
        Write-Log "Error checking duplicate SUG deployment: $_" -Level ERROR
        return [PSCustomObject]@{ DuplicateCheckFailed = $true; Error = $_.Exception.Message }
    }
}

# ---------------------------------------------------------------------------
# Task Sequences
# ---------------------------------------------------------------------------

function Search-CMTaskSequenceByName {
    param([Parameter(Mandatory)][string]$SearchText)

    try {
        $tsList = @(Get-CMTaskSequence -Name "*$SearchText*" -Fast -ErrorAction Stop |
            Select-Object Name, PackageID, BootImageID, Description |
            Sort-Object Name)
        Write-Log "Task sequence search '$SearchText': $($tsList.Count) result(s)"
        return $tsList
    }
    catch {
        Write-Log "Error searching task sequences: $_" -Level ERROR
        return @()
    }
}

function Test-TaskSequenceExists {
    param([Parameter(Mandatory)][string]$TaskSequenceName)

    try {
        $ts = Get-CMTaskSequence -Name $TaskSequenceName -Fast -DisableWildcardHandling -ErrorAction Stop
        if ($null -eq $ts) {
            Write-Log "Task sequence not found: $TaskSequenceName" -Level WARN
            return $null
        }
        Write-Log "Task sequence found: $TaskSequenceName (PackageID: $($ts.PackageID))"
        return $ts
    }
    catch {
        Write-Log "Error querying task sequence '$TaskSequenceName': $_" -Level ERROR
        return $null
    }
}

function Test-DuplicateTaskSequenceDeployment {
    param(
        [Parameter(Mandatory)][string]$TaskSequencePackageId,
        [Parameter(Mandatory)][string]$CollectionName
    )

    try {
        $existing = Get-CMTaskSequenceDeployment -TaskSequenceId $TaskSequencePackageId -CollectionName $CollectionName -Fast -DisableWildcardHandling -ErrorAction Stop
        if ($null -ne $existing -and @($existing).Count -gt 0) {
            Write-Log "Duplicate TS deployment: TaskSequencePackageId=$TaskSequencePackageId Collection='$CollectionName'" -Level WARN
            return $existing
        }
        Write-Log "No duplicate TS deployment: TaskSequencePackageId=$TaskSequencePackageId Collection='$CollectionName'"
        return $null
    }
    catch {
        Write-Log "Error checking duplicate TS deployment: $_" -Level ERROR
        return [PSCustomObject]@{ DuplicateCheckFailed = $true; Error = $_.Exception.Message }
    }
}

function Invoke-TaskSequenceDeployment {
    <#
    .SYNOPSIS
        Deploys a task sequence to a device collection using New-CMTaskSequenceDeployment.

    .DESCRIPTION
        Wraps New-CMTaskSequenceDeployment with -InputObject.
        UTC flag drives both UseUtcForAvailableSchedule and UseUtcForExpireSchedule
        so the available + expire times stay in the same zone.
        The Required deadline goes to -Schedule; the cmdlet's -DeadlineDateTime
        is the expiry, which this function never sets.
    #>
    param(
        [Parameter(Mandatory)]$TaskSequence,
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)][ValidateSet('Required','Available')][string]$DeployPurpose,
        [Parameter(Mandatory)][datetime]$AvailableDateTime,
        [datetime]$DeadlineDateTime,
        [ValidateSet('Clients','ClientsMediaAndPxe','MediaAndPxe','MediaAndPxeHidden')]
        [string]$Availability = 'Clients',
        [ValidateSet('LocalTime','Utc')][string]$TimeBasedOn = 'LocalTime',
        [bool]$ShowTaskSequenceProgress = $true,
        [bool]$OverrideServiceWindow = $false,
        [bool]$RebootOutsideServiceWindow = $false,
        [bool]$UseMeteredNetwork = $false
    )

    $blocked = Get-UnsafeCollectionResult -Collection $Collection
    if ($blocked) { return $blocked }

    $params = @{
        InputObject               = $TaskSequence
        CollectionId              = $Collection.CollectionID
        DeployPurpose             = $DeployPurpose
        AvailableDateTime         = $AvailableDateTime
        Availability              = $Availability
        ShowTaskSequenceProgress  = $ShowTaskSequenceProgress
        ErrorAction               = 'Stop'
    }

    if ($TimeBasedOn -eq 'Utc') {
        $params['UseUtcForAvailableSchedule'] = $true
        $params['UseUtcForExpireSchedule']    = $true
    }
    if ($OverrideServiceWindow)      { $params['SoftwareInstallation'] = $true }
    if ($RebootOutsideServiceWindow) { $params['SystemRestart']        = $true }
    # UseMeteredNetwork is a Required-only param. Passing it on Available
    # TS deploys triggers a cmdlet WARN that leaks into the audit log
    # (ConfigMgr surfaces "Parameter X does not apply to deployments with Purpose
    # Available"). Apps + Packages already model this gating via their
    # $params hashtable conditions; mirror it here.
    if ($DeployPurpose -eq 'Required' -and $UseMeteredNetwork) {
        $params['UseMeteredNetwork'] = $true
    }

    try {
        if ($DeployPurpose -eq 'Required' -and $DeadlineDateTime) {
            $params['Schedule'] = @(New-RequiredAssignmentSchedule -Deadline $DeadlineDateTime -TimeBasedOn $TimeBasedOn)
        }

        Write-Log ("Executing TS deployment: {0} -> {1} ({2} devices) as {3}, Availability={4}" -f
            $TaskSequence.Name, $Collection.Name, $Collection.MemberCount, $DeployPurpose, $Availability)

        $deployment = New-CMTaskSequenceDeployment @params

        # Same fix as packages: New-CMTaskSequenceDeployment returns
        # SMS_Advertisement -- AdvertisementID, not DeploymentID.
        Write-Log "TS deployment created (AdvertisementID: $($deployment.AdvertisementID))"
        return @{
            Success            = $true
            DeploymentID       = $deployment.AdvertisementID
            DeploymentUniqueID = [string]$deployment.AdvertisementID
            Error              = $null
        }
    }
    catch {
        Write-Log "TS deployment FAILED: $_" -Level ERROR
        return @{
            Success            = $false
            DeploymentID       = $null
            DeploymentUniqueID = $null
            Error              = $_.ToString()
        }
    }
}

function New-RequiredAssignmentSchedule {
    <#
    .SYNOPSIS
        Returns the one-time schedule token that New-CMPackageDeployment and
        New-CMTaskSequenceDeployment take as -Schedule (the Required deadline).
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Returns a schedule token object; the site is not changed.')]
    param(
        [Parameter(Mandatory)][datetime]$Deadline,
        [ValidateSet('LocalTime','Utc')][string]$TimeBasedOn = 'LocalTime'
    )

    $p = @{ Start = $Deadline; Nonrecurring = $true; ErrorAction = 'Stop' }
    if ($TimeBasedOn -eq 'Utc') { $p['IsUtc'] = $true }
    return New-CMSchedule @p
}

function Invoke-PackageDeployment {
    <#
    .SYNOPSIS
        Deploys a legacy package + program to a device collection using New-CMPackageDeployment.

    .DESCRIPTION
        Wraps New-CMPackageDeployment with the standard-program parameter set.
        Success/failure is returned as a hashtable with DeploymentID.

        UTC semantics: TimeBasedOn='Utc' sets both -UseUtcForAvailableSchedule and
        -UseUtcForExpireSchedule. The package cmdlet treats these independently in ConfigMgr;
        for safety-critical tool usage we keep them in lockstep.

        The Required deadline goes to -Schedule; the cmdlet's -DeadlineDateTime
        is the expiry, which this function never sets.

        Maintenance-window semantics:
          -OverrideServiceWindow maps to -SoftwareInstallation $true (install outside MW)
          -RebootOutsideServiceWindow maps to -SystemRestart $true
    #>
    param(
        [Parameter(Mandatory)]$Package,
        [Parameter(Mandatory)][string]$ProgramName,
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)][ValidateSet('Required','Available')][string]$DeployPurpose,
        [Parameter(Mandatory)][datetime]$AvailableDateTime,
        [datetime]$DeadlineDateTime,
        [ValidateSet('LocalTime','Utc')][string]$TimeBasedOn = 'LocalTime',
        [bool]$OverrideServiceWindow = $false,
        [bool]$RebootOutsideServiceWindow = $false,
        [bool]$UseMeteredNetwork = $false,
        # MS enum asymmetry: Fast has "AndRunLocally", Slow has "AndLocally"
        # (no "Run"). Slow also uses "FromDistributionPoint", not
        # "FromRemoteDistributionPoint". Do not "normalize" these -- the
        # cmdlet rejects any other spelling with a cryptic enum-bind error.
        [ValidateSet('DownloadContentFromDistributionPointAndRunLocally','RunProgramFromDistributionPoint')]
        [string]$FastNetworkOption = 'DownloadContentFromDistributionPointAndRunLocally',
        [ValidateSet('DoNotRunProgram','DownloadContentFromDistributionPointAndLocally','RunProgramFromDistributionPoint')]
        [string]$SlowNetworkOption = 'DoNotRunProgram',
        [ValidateSet('NeverRerunDeployedProgram','AlwaysRerunProgram','RerunIfFailedPreviousAttempt','RerunIfSucceededOnPreviousAttempt')]
        [string]$RerunBehavior = 'NeverRerunDeployedProgram'
    )

    $blocked = Get-UnsafeCollectionResult -Collection $Collection
    if ($blocked) { return $blocked }

    $params = @{
        StandardProgram   = $true
        PackageId         = $Package.PackageID
        ProgramName       = $ProgramName
        CollectionId      = $Collection.CollectionID
        DeployPurpose     = $DeployPurpose
        AvailableDateTime = $AvailableDateTime
        FastNetworkOption = $FastNetworkOption
        SlowNetworkOption = $SlowNetworkOption
        RerunBehavior     = $RerunBehavior
        ErrorAction       = 'Stop'
    }

    if ($TimeBasedOn -eq 'Utc') {
        $params['UseUtcForAvailableSchedule'] = $true
        $params['UseUtcForExpireSchedule']    = $true
    }
    if ($OverrideServiceWindow)      { $params['SoftwareInstallation'] = $true }
    if ($RebootOutsideServiceWindow) { $params['SystemRestart']        = $true }
    if ($UseMeteredNetwork)          { $params['UseMeteredNetwork']    = $true }

    try {
        if ($DeployPurpose -eq 'Required' -and $DeadlineDateTime) {
            $params['Schedule'] = @(New-RequiredAssignmentSchedule -Deadline $DeadlineDateTime -TimeBasedOn $TimeBasedOn)
        }

        Write-Log ("Executing package deployment: {0} / {1} -> {2} ({3} devices) as {4}" -f
            $Package.Name, $ProgramName, $Collection.Name, $Collection.MemberCount, $DeployPurpose)

        $deployment = New-CMPackageDeployment @params

        # New-CMPackageDeployment returns SMS_Advertisement (legacy pkg
        # deployment shape). The ID property is AdvertisementID; the
        # earlier $deployment.DeploymentID read was returning $null.
        Write-Log "Package deployment created (AdvertisementID: $($deployment.AdvertisementID))"
        return @{
            Success            = $true
            DeploymentID       = $deployment.AdvertisementID
            DeploymentUniqueID = [string]$deployment.AdvertisementID
            Error              = $null
        }
    }
    catch {
        Write-Log "Package deployment FAILED: $_" -Level ERROR
        return @{
            Success            = $false
            DeploymentID       = $null
            DeploymentUniqueID = $null
            Error              = $_.ToString()
        }
    }
}

# ---------------------------------------------------------------------------
# Ring plans
# ---------------------------------------------------------------------------
# A ring plan is data: Rings\<plan>.json. Loading validates the plan and
# refuses one that names a built-in collection. Expansion turns day offsets
# into absolute dates for one start time; nothing here calls the site.

$script:RingDateFormat     = 'yyyy-MM-ddTHH:mm:ss'
$script:RingNotification   = @('DisplayAll', 'DisplaySoftwareCenterOnly', 'HideAll')
$script:RingTsAvailability = @('Clients', 'ClientsMediaAndPxe', 'MediaAndPxe', 'MediaAndPxeHidden')
$script:RingFastNetwork    = @('DownloadContentFromDistributionPointAndRunLocally', 'RunProgramFromDistributionPoint')
$script:RingSlowNetwork    = @('DoNotRunProgram', 'DownloadContentFromDistributionPointAndLocally', 'RunProgramFromDistributionPoint')
$script:RingRerun          = @('NeverRerunDeployedProgram', 'AlwaysRerunProgram', 'RerunIfFailedPreviousAttempt', 'RerunIfSucceededOnPreviousAttempt')

function ConvertFrom-RingDateText {
    <#
    .SYNOPSIS
        Parses a ring date. Accepts a DateTime or invariant text
        (yyyy-MM-ddTHH:mm:ss, yyyy-MM-dd HH:mm:ss, yyyy-MM-dd HH:mm).
        Returns $null for empty input and throws on anything else.
    #>
    param([AllowNull()][AllowEmptyString()]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return [datetime]$Value }
    $s = ([string]$Value).Trim()
    if ($s.Length -eq 0) { return $null }
    $formats = [string[]]@('yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-ddTHH:mm', 'yyyy-MM-dd HH:mm')
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($s, $formats, [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
        return $d
    }
    throw "Date '$s' is not in the form yyyy-MM-dd HH:mm."
}

function ConvertTo-RingDateText {
    param([AllowNull()][AllowEmptyString()]$Value)
    $d = ConvertFrom-RingDateText -Value $Value
    if ($null -eq $d) { return $null }
    return $d.ToString($script:RingDateFormat, [System.Globalization.CultureInfo]::InvariantCulture)
}

function ConvertTo-RingNumber {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string] -and $Value.Trim().Length -eq 0) { return $null }
    $n = 0.0
    if ([double]::TryParse([string]$Value, [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
        return $n
    }
    return [double]::NaN
}

function Get-RingPropertyValue {
    param($InputObject, [Parameter(Mandatory)][string]$Name, $Default)
    if ($null -eq $InputObject) { return $Default }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) { return $InputObject[$Name] }
        return $Default
    }
    $p = $InputObject.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $Default
}

function Get-RingPlanSeed {
    <#
    .SYNOPSIS
        Returns the two plans written on first run. Every CollectionID is
        empty so a seed cannot deploy anywhere until an operator sets targets.
    #>
    $ring = {
        param($Name, $Avail, $Dead, $Notify, $Threshold)
        [ordered]@{
            Name                       = $Name
            CollectionID               = ''
            CollectionName             = ''
            Purpose                    = 'Required'
            AvailableOffsetDays        = $Avail
            DeadlineOffsetDays         = $Dead
            UserNotification           = $Notify
            OverrideServiceWindow      = $false
            RebootOutsideServiceWindow = $false
            AllowMeteredConnection     = $false
            DPGroup                    = ''
            SuccessThresholdPercent    = $Threshold
        }
    }
    return @(
        [ordered]@{
            Name           = 'Workstation-Rings'
            Description    = 'Four workstation rings. Set each CollectionID before use.'
            HoldLaterRings = $false
            TimeBasedOn    = 'LocalTime'
            Rings          = @(
                (& $ring 'QA'         0  1 'DisplayAll' $null),
                (& $ring 'Pilot'      2  5 'DisplayAll' 90),
                (& $ring 'Prod 1'     7 10 'DisplayAll' 95),
                (& $ring 'Prod Final' 14 17 'DisplayAll' $null)
            )
        },
        [ordered]@{
            Name           = 'Server-Rings'
            Description    = 'Two server rings; Prod waits for Promote. Set each CollectionID before use.'
            HoldLaterRings = $true
            TimeBasedOn    = 'LocalTime'
            Rings          = @(
                (& $ring 'Test' 0 2 'HideAll' 95),
                (& $ring 'Prod' 7 9 'HideAll' $null)
            )
        }
    )
}

function Initialize-RingPlanFolder {
    <#
    .SYNOPSIS
        Creates the plan folder and writes the seed plans when it holds no
        *.json file. Existing plans are never overwritten. Returns the number
        of seed files written.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='First-run seed of an empty local folder; never overwrites.')]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    if (Get-ChildItem -LiteralPath $Path -Filter '*.json' -File -ErrorAction SilentlyContinue) { return 0 }

    $written = 0
    foreach ($plan in Get-RingPlanSeed) {
        $file = Join-Path $Path ($plan.Name + '.json')
        $plan | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $file -Encoding UTF8
        $written++
    }
    Write-Log "Seeded $written ring plan(s) in $Path"
    return $written
}

function ConvertTo-RingPlan {
    <#
    .SYNOPSIS
        Normalizes a plan read from JSON (or a hashtable) and validates it.
        Returns @{ Plan; Errors }. Missing ring options take the defaults the
        single-deployment flow uses.
    #>
    param([Parameter(Mandatory)][AllowNull()]$InputObject)

    $errors = New-Object System.Collections.Generic.List[string]
    if ($null -eq $InputObject) {
        $errors.Add('The plan file is empty.')
        return @{ Plan = $null; Errors = @($errors) }
    }

    $rings = @()
    $i = 0
    foreach ($r in @(Get-RingPropertyValue -InputObject $InputObject -Name 'Rings' -Default @())) {
        if ($null -eq $r) { continue }
        $i++
        $rings += [PSCustomObject][ordered]@{
            Index                      = $i
            Name                       = ([string](Get-RingPropertyValue -InputObject $r -Name 'Name' -Default '')).Trim()
            CollectionID               = ([string](Get-RingPropertyValue -InputObject $r -Name 'CollectionID' -Default '')).Trim().ToUpperInvariant()
            CollectionName             = ([string](Get-RingPropertyValue -InputObject $r -Name 'CollectionName' -Default '')).Trim()
            Purpose                    = [string](Get-RingPropertyValue -InputObject $r -Name 'Purpose' -Default 'Required')
            AvailableOffsetDays        = ConvertTo-RingNumber (Get-RingPropertyValue -InputObject $r -Name 'AvailableOffsetDays' -Default 0)
            DeadlineOffsetDays         = ConvertTo-RingNumber (Get-RingPropertyValue -InputObject $r -Name 'DeadlineOffsetDays' -Default $null)
            UserNotification           = [string](Get-RingPropertyValue -InputObject $r -Name 'UserNotification' -Default 'DisplayAll')
            OverrideServiceWindow      = [bool](Get-RingPropertyValue -InputObject $r -Name 'OverrideServiceWindow' -Default $false)
            RebootOutsideServiceWindow = [bool](Get-RingPropertyValue -InputObject $r -Name 'RebootOutsideServiceWindow' -Default $false)
            AllowMeteredConnection     = [bool](Get-RingPropertyValue -InputObject $r -Name 'AllowMeteredConnection' -Default $false)
            AllowBoundaryFallback      = [bool](Get-RingPropertyValue -InputObject $r -Name 'AllowBoundaryFallback' -Default $true)
            AllowMicrosoftUpdate       = [bool](Get-RingPropertyValue -InputObject $r -Name 'AllowMicrosoftUpdate' -Default $false)
            RequirePostRebootFullScan  = [bool](Get-RingPropertyValue -InputObject $r -Name 'RequirePostRebootFullScan' -Default $true)
            ShowTaskSequenceProgress   = [bool](Get-RingPropertyValue -InputObject $r -Name 'ShowTaskSequenceProgress' -Default $true)
            TaskSequenceAvailability   = [string](Get-RingPropertyValue -InputObject $r -Name 'TaskSequenceAvailability' -Default 'Clients')
            FastNetworkOption          = [string](Get-RingPropertyValue -InputObject $r -Name 'FastNetworkOption' -Default 'DownloadContentFromDistributionPointAndRunLocally')
            SlowNetworkOption          = [string](Get-RingPropertyValue -InputObject $r -Name 'SlowNetworkOption' -Default 'DoNotRunProgram')
            RerunBehavior              = [string](Get-RingPropertyValue -InputObject $r -Name 'RerunBehavior' -Default 'NeverRerunDeployedProgram')
            DPGroup                    = ([string](Get-RingPropertyValue -InputObject $r -Name 'DPGroup' -Default '')).Trim()
            SuccessThresholdPercent    = ConvertTo-RingNumber (Get-RingPropertyValue -InputObject $r -Name 'SuccessThresholdPercent' -Default $null)
        }
    }

    $plan = [PSCustomObject][ordered]@{
        Name           = ([string](Get-RingPropertyValue -InputObject $InputObject -Name 'Name' -Default '')).Trim()
        Description    = [string](Get-RingPropertyValue -InputObject $InputObject -Name 'Description' -Default '')
        HoldLaterRings = [bool](Get-RingPropertyValue -InputObject $InputObject -Name 'HoldLaterRings' -Default $false)
        TimeBasedOn    = [string](Get-RingPropertyValue -InputObject $InputObject -Name 'TimeBasedOn' -Default 'LocalTime')
        Rings          = $rings
        FilePath       = ''
    }

    foreach ($e in @(Test-RingPlan -Plan $plan)) { $errors.Add($e) }
    return @{ Plan = $plan; Errors = @($errors) }
}

function Get-RingLabel {
    param([Parameter(Mandatory)]$Ring)
    return ("Ring {0} '{1}'" -f $Ring.Index, $Ring.Name)
}

function Test-RingPlan {
    <#
    .SYNOPSIS
        Returns the list of problems in a normalized plan. An empty list means
        the plan loads. An empty CollectionID is allowed here (seed plans);
        Test-RingExpansion refuses it before anything runs.
    #>
    param([Parameter(Mandatory)]$Plan)

    $errors = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Plan.Name)) { $errors.Add('The plan has no Name.') }
    if ($Plan.TimeBasedOn -notin @('LocalTime', 'Utc')) {
        $errors.Add(("TimeBasedOn '{0}' is not LocalTime or Utc." -f $Plan.TimeBasedOn))
    }
    $rings = @($Plan.Rings)
    if ($rings.Count -eq 0) { $errors.Add('The plan has no rings.') }

    $seen = @{}
    $lastDeadline = $null
    $lastDeadlineLabel = ''
    foreach ($r in $rings) {
        $label = Get-RingLabel -Ring $r
        if ([string]::IsNullOrWhiteSpace($r.Name)) { $errors.Add(("Ring {0} has no Name." -f $r.Index)) }

        if ($r.CollectionID) {
            if (Test-CollectionIdBuiltIn -CollectionId $r.CollectionID) {
                $errors.Add(("{0} targets built-in collection {1}. Built-in collections are blocked." -f $label, $r.CollectionID))
            }
            elseif ($r.CollectionID -notmatch '^[A-Z0-9]{8}$') {
                $errors.Add(("{0} CollectionID '{1}' is not an 8-character collection ID." -f $label, $r.CollectionID))
            }
            if ($seen.ContainsKey($r.CollectionID)) {
                $errors.Add(("{0} targets {1}, which {2} also targets. Rings must not share a collection." -f $label, $r.CollectionID, $seen[$r.CollectionID]))
            }
            else { $seen[$r.CollectionID] = $label }
        }

        $availOk = ($null -ne $r.AvailableOffsetDays -and -not [double]::IsNaN($r.AvailableOffsetDays) -and $r.AvailableOffsetDays -ge 0)
        if ($r.Purpose -notin @('Available', 'Required')) {
            $errors.Add(("{0} Purpose '{1}' is not Available or Required." -f $label, $r.Purpose))
        }
        if (-not $availOk) {
            $errors.Add(("{0} AvailableOffsetDays must be a number of days, 0 or more." -f $label))
        }
        if ($r.Purpose -eq 'Required') {
            if ($null -eq $r.DeadlineOffsetDays -or [double]::IsNaN($r.DeadlineOffsetDays) -or $r.DeadlineOffsetDays -lt 0) {
                $errors.Add(("{0} is Required and needs DeadlineOffsetDays, 0 or more." -f $label))
            }
            else {
                if ($availOk -and $r.DeadlineOffsetDays -le $r.AvailableOffsetDays) {
                    $errors.Add(("{0} deadline offset ({1}) must be later than its available offset ({2})." -f $label, $r.DeadlineOffsetDays, $r.AvailableOffsetDays))
                }
                if ($null -ne $lastDeadline -and $r.DeadlineOffsetDays -le $lastDeadline) {
                    $errors.Add(("{0} deadline offset ({1}) must be later than the {2} deadline offset ({3})." -f $label, $r.DeadlineOffsetDays, $lastDeadlineLabel, $lastDeadline))
                }
                $lastDeadline = $r.DeadlineOffsetDays
                $lastDeadlineLabel = $label
            }
        }

        if ($r.UserNotification -notin $script:RingNotification) {
            $errors.Add(("{0} UserNotification '{1}' is not valid." -f $label, $r.UserNotification))
        }
        elseif ($r.Purpose -eq 'Available' -and $r.UserNotification -eq 'HideAll') {
            $errors.Add(("{0} is Available with HideAll: users never see it and it never installs." -f $label))
        }
        if ($r.TaskSequenceAvailability -notin $script:RingTsAvailability) {
            $errors.Add(("{0} TaskSequenceAvailability '{1}' is not valid." -f $label, $r.TaskSequenceAvailability))
        }
        if ($r.FastNetworkOption -notin $script:RingFastNetwork) {
            $errors.Add(("{0} FastNetworkOption '{1}' is not valid." -f $label, $r.FastNetworkOption))
        }
        if ($r.SlowNetworkOption -notin $script:RingSlowNetwork) {
            $errors.Add(("{0} SlowNetworkOption '{1}' is not valid." -f $label, $r.SlowNetworkOption))
        }
        if ($r.RerunBehavior -notin $script:RingRerun) {
            $errors.Add(("{0} RerunBehavior '{1}' is not valid." -f $label, $r.RerunBehavior))
        }
        if ($null -ne $r.SuccessThresholdPercent -and
            ([double]::IsNaN($r.SuccessThresholdPercent) -or $r.SuccessThresholdPercent -lt 0 -or $r.SuccessThresholdPercent -gt 100)) {
            $errors.Add(("{0} SuccessThresholdPercent must be empty or 0 to 100." -f $label))
        }
    }
    return @($errors)
}

function Import-RingPlan {
    <#
    .SYNOPSIS
        Loads one plan file. Throws when the file does not parse or the plan
        fails validation; the message names each failing ring.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $leaf = Split-Path -Path $Path -Leaf
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 -ErrorAction Stop
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw ("Plan {0} failed to load: {1}" -f $leaf, $_.Exception.Message)
    }
    $result = ConvertTo-RingPlan -InputObject $obj
    if (@($result.Errors).Count -gt 0) {
        throw ("Plan {0} failed to load:`n- {1}" -f $leaf, (@($result.Errors) -join "`n- "))
    }
    $result.Plan.FilePath = $Path
    return $result.Plan
}

function Get-RingPlanList {
    <#
    .SYNOPSIS
        Loads every *.json plan in a folder. Returns @{ Plans; Errors }; a plan
        that fails to load is left out and its error returned.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $plans  = @()
    $errors = @()
    if (-not (Test-Path -LiteralPath $Path)) { return @{ Plans = @(); Errors = @() } }
    foreach ($f in @(Get-ChildItem -LiteralPath $Path -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        try { $plans += Import-RingPlan -Path $f.FullName }
        catch {
            $errors += $_.Exception.Message
            Write-Log $_.Exception.Message -Level WARN
        }
    }
    Write-Log ("Ring plans: {0} loaded, {1} refused" -f $plans.Count, $errors.Count)
    return @{ Plans = $plans; Errors = $errors }
}

function Expand-RingPlan {
    <#
    .SYNOPSIS
        Expands a plan into one row per ring with absolute dates. Pure: the
        start time is the only clock input.
    #>
    param(
        [Parameter(Mandatory)]$Plan,
        [Parameter(Mandatory)][datetime]$Start
    )

    $rows = foreach ($r in @($Plan.Rings)) {
        $row = [ordered]@{}
        foreach ($p in $r.PSObject.Properties) { $row[$p.Name] = $p.Value }
        $row['AvailableDateTime'] = $Start.AddDays([double]$r.AvailableOffsetDays)
        $row['DeadlineDateTime']  = if ($r.Purpose -eq 'Required' -and $null -ne $r.DeadlineOffsetDays) { $Start.AddDays([double]$r.DeadlineOffsetDays) } else { $null }
        [PSCustomObject]$row
    }
    return ,@($rows)
}

function Get-RingNow {
    <#
    .SYNOPSIS
        The clock that ring times are compared with. Ring times of a plan with
        TimeBasedOn Utc are UTC wall-clock values, so they compare with
        UtcNow; DateTime comparison ignores Kind and compares ticks only.
    #>
    param([ValidateSet('LocalTime', 'Utc')][string]$TimeBasedOn = 'LocalTime')
    if ($TimeBasedOn -eq 'Utc') { return [datetime]::UtcNow }
    return Get-Date
}

function Test-RingExpansion {
    <#
    .SYNOPSIS
        Validates the rows a ring run submits. Returns @{ Errors; Warnings }.
        Errors block the run; warnings need the operator to confirm.
    #>
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Rings,
        [datetime]$Now = (Get-Date)
    )

    $errors   = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    if (@($Rings).Count -eq 0) { $errors.Add('There are no rings to run.') }

    $seen = @{}
    $lastDeadline = $null
    $lastLabel = ''
    foreach ($r in @($Rings)) {
        $label = Get-RingLabel -Ring $r
        $cid = ([string]$r.CollectionID).Trim().ToUpperInvariant()
        if ($cid.Length -eq 0) {
            $errors.Add(("{0} has no target collection." -f $label))
        }
        elseif (Test-CollectionIdBuiltIn -CollectionId $cid) {
            $errors.Add(("{0} targets built-in collection {1}. Built-in collections are blocked." -f $label, $cid))
        }
        elseif ($seen.ContainsKey($cid)) {
            $errors.Add(("{0} targets {1}, which {2} also targets. Rings must not share a collection." -f $label, $cid, $seen[$cid]))
        }
        else { $seen[$cid] = $label }

        $available = ConvertFrom-RingDateText $r.AvailableDateTime
        $deadline  = ConvertFrom-RingDateText $r.DeadlineDateTime
        if ($null -eq $available) { $errors.Add(("{0} has no available time." -f $label)); continue }

        if ($r.Purpose -eq 'Required') {
            if ($null -eq $deadline) {
                $errors.Add(("{0} is Required and has no deadline." -f $label))
            }
            else {
                if ($deadline -le $available) {
                    $errors.Add(("{0} deadline {1} must be later than its available time {2}." -f $label,
                        $deadline.ToString('yyyy-MM-dd HH:mm'), $available.ToString('yyyy-MM-dd HH:mm')))
                }
                if ($deadline -le $Now) {
                    $errors.Add(("{0} deadline {1} is in the past; the deployment would enforce at once." -f $label, $deadline.ToString('yyyy-MM-dd HH:mm')))
                }
                if ($null -ne $lastDeadline -and $deadline -le $lastDeadline) {
                    $errors.Add(("{0} deadline {1} must be later than the {2} deadline {3}." -f $label,
                        $deadline.ToString('yyyy-MM-dd HH:mm'), $lastLabel, $lastDeadline.ToString('yyyy-MM-dd HH:mm')))
                }
                $lastDeadline = $deadline
                $lastLabel = $label
            }
        }
        if ($r.Purpose -eq 'Available' -and $r.UserNotification -eq 'HideAll') {
            $errors.Add(("{0} is Available with HideAll: users never see it and it never installs." -f $label))
        }
        if ($available -lt $Now.AddMinutes(-5)) {
            $warnings.Add(("{0} available time {1} is in the past; clients see it at once." -f $label, $available.ToString('yyyy-MM-dd HH:mm')))
        }
    }
    return @{ Errors = @($errors); Warnings = @($warnings) }
}

# ---------------------------------------------------------------------------
# Ring runs (hold later rings)
# ---------------------------------------------------------------------------
# Run file: <plan>_<objectKey>_<yyyyMMdd-HHmmss>.json in the run-state folder.
# Ring status moves Held -> Creating -> Created and Created -> Removed
# (Reconcile). A finished run moves to the closed subfolder.

function New-RingRunId {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Returns a new GUID string.')]
    param()
    return [guid]::NewGuid().ToString()
}

function Get-RingSafeFileToken {
    param([AllowEmptyString()][AllowNull()][string]$Text)
    $s = ([string]$Text) -replace '[^A-Za-z0-9_\-]', '_'
    if ([string]::IsNullOrWhiteSpace($s)) { $s = 'unnamed' }
    return $s
}

function Get-RingRunFileName {
    param(
        [Parameter(Mandatory)][string]$PlanName,
        [Parameter(Mandatory)][string]$ObjectKey,
        [Parameter(Mandatory)][datetime]$Stamp
    )
    return ('{0}_{1}_{2}.json' -f (Get-RingSafeFileToken $PlanName), (Get-RingSafeFileToken $ObjectKey), $Stamp.ToString('yyyyMMdd-HHmmss'))
}

function Get-RingFreePath {
    <#
    .SYNOPSIS
        Returns Path when no file exists there, else the first free
        <name>-2.json, <name>-3.json, ... beside it. With -AlsoCheckFolder, a
        name taken in that folder counts as taken too.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$AlsoCheckFolder
    )
    $dir  = Split-Path -Path $Path -Parent
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $ext  = [System.IO.Path]::GetExtension($Path)
    $n = 1
    $candidate = $Path
    while ((Test-Path -LiteralPath $candidate) -or
           ($AlsoCheckFolder -and (Test-Path -LiteralPath (Join-Path $AlsoCheckFolder (Split-Path -Path $candidate -Leaf))))) {
        $n++
        $candidate = Join-Path $dir ('{0}-{1}{2}' -f $base, $n, $ext)
    }
    return $candidate
}

function New-RingRunFile {
    <#
    .SYNOPSIS
        Writes a new run file under a name no open or closed run uses. The
        file is opened with FileMode.CreateNew, so two sessions that pick the
        same name at the same moment cannot both write it. Returns the path.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Creates the tool''s own run-state file.')]
    param(
        [Parameter(Mandatory)]$Run,
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$FileName
    )

    if (-not (Test-Path -LiteralPath $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
    $closed = Join-Path $Folder 'closed'
    $json = $Run | ConvertTo-Json -Depth 8
    $bytes = (New-Object System.Text.UTF8Encoding $false).GetBytes($json)
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        $path = Get-RingFreePath -Path (Join-Path $Folder $FileName) -AlsoCheckFolder $closed
        try {
            $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        }
        catch {
            if (Test-Path -LiteralPath $path) { continue }
            throw
        }
        try { $fs.Write($bytes, 0, $bytes.Length) } finally { $fs.Dispose() }
        return $path
    }
    throw "No free run file name for $FileName in $Folder."
}

function Get-RingObjectIdentity {
    param(
        [Parameter(Mandatory)][ValidateSet('Application', 'Package', 'TaskSequence', 'SUG')][string]$Type,
        [Parameter(Mandatory)]$TargetObject,
        [string]$ProgramName
    )
    switch ($Type) {
        'Application'  { $id = [string]$TargetObject.PackageID; $name = [string]$TargetObject.LocalizedDisplayName }
        'Package'      { $id = [string]$TargetObject.PackageID; $name = [string]$TargetObject.Name }
        'TaskSequence' { $id = [string]$TargetObject.PackageID; $name = [string]$TargetObject.Name }
        'SUG'          { $id = [string]$TargetObject.CI_ID;     $name = [string]$TargetObject.LocalizedDisplayName }
    }
    return [ordered]@{
        Type        = $Type
        ID          = $id
        Name        = $name
        ProgramName = if ($Type -eq 'Package') { $ProgramName } else { $null }
    }
}

function Get-DuplicateCheckFailure {
    param([AllowNull()]$Result)

    foreach ($candidate in @($Result)) {
        if ($null -eq $candidate) { continue }
        $marker = $candidate.PSObject.Properties['DuplicateCheckFailed']
        if ($null -ne $marker -and [bool]$marker.Value) { return $candidate }
    }
    return $null
}

function Get-RingDeploymentReference {
    param(
        [Parameter(Mandatory)][ValidateSet('Application', 'Package', 'TaskSequence', 'SUG')][string]$Type,
        [Parameter(Mandatory)]$Deployment
    )

    $idProperties = switch ($Type) {
        'Application'  { @('AssignmentUniqueID', 'DeploymentUniqueID', 'DeploymentID') }
        'Package'      { @('AdvertisementID', 'DeploymentID') }
        'TaskSequence' { @('AdvertisementID', 'DeploymentID') }
        'SUG'          { @('AssignmentUniqueID', 'DeploymentUniqueID', 'DeploymentID') }
    }
    $deploymentId = ''
    foreach ($propertyName in $idProperties) {
        $property = $Deployment.PSObject.Properties[$propertyName]
        if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            $deploymentId = [string]$property.Value
            break
        }
    }
    $assignmentId = $null
    $assignmentProperty = $Deployment.PSObject.Properties['AssignmentID']
    if ($null -ne $assignmentProperty -and $null -ne $assignmentProperty.Value) {
        $assignmentId = [string]$assignmentProperty.Value
    }
    if ([string]::IsNullOrWhiteSpace($deploymentId)) {
        throw ("Could not determine the deployment ID for a {0} deployment during recovery." -f $Type)
    }
    return @{ DeploymentID = $deploymentId; AssignmentID = $assignmentId }
}

function New-RingRun {
    <#
    .SYNOPSIS
        Builds the run-state object for a hold-later-rings run. Every ring
        starts Held; the caller creates ring 1 and records it.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Builds an in-memory run object.')]
    param(
        [Parameter(Mandatory)][string]$PlanName,
        [Parameter(Mandatory)][System.Collections.IDictionary]$Object,
        [Parameter(Mandatory)][array]$Rings,
        [Parameter(Mandatory)][string]$RunId,
        [ValidateSet('LocalTime', 'Utc')][string]$TimeBasedOn = 'LocalTime',
        [datetime]$CreatedAt = (Get-Date)
    )

    # Display-only columns of the preview grid; the run file keeps data only.
    $skip = @('AvailableText', 'DeadlineText', 'ThresholdText', 'Checks')
    $ringRows = foreach ($r in $Rings) {
        $row = [ordered]@{}
        foreach ($p in $r.PSObject.Properties) { if ($p.Name -notin $skip) { $row[$p.Name] = $p.Value } }
        $row['AvailableDateTime'] = ConvertTo-RingDateText $r.AvailableDateTime
        $row['DeadlineDateTime']  = ConvertTo-RingDateText $r.DeadlineDateTime
        $row['Status']            = 'Held'
        $row['DeploymentID']      = $null
        $row['AssignmentID']      = $null
        $row['CreatingAt']        = $null
        $row['CreatedAt']         = $null
        $row['RemovedAt']         = $null
        $row['ShiftedMinutes']    = 0
        $row
    }

    $run = [ordered]@{
        SchemaVersion  = 1
        RunId          = $RunId
        PlanName       = $PlanName
        TimeBasedOn    = $TimeBasedOn
        HoldLaterRings = $true
        CreatedAt      = ConvertTo-RingDateText $CreatedAt
        CreatedBy      = "$env:USERDOMAIN\$env:USERNAME"
        ClosedAt       = $null
        Object         = $Object
        Rings          = @($ringRows)
    }
    $copy = $run | ConvertTo-Json -Depth 8 | ConvertFrom-Json
    $copy.Rings = @($copy.Rings)
    return $copy
}

function Save-RingRun {
    <#
    .SYNOPSIS
        Writes the run file through a temporary file and a replace, so a
        reader never sees a half-written file.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Writes the tool''s own run-state file.')]
    param(
        [Parameter(Mandatory)]$Run,
        [Parameter(Mandatory)][string]$Path
    )

    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $json = $Run | ConvertTo-Json -Depth 8
    $tmp  = $Path + '.tmp'
    [System.IO.File]::WriteAllText($tmp, $json, (New-Object System.Text.UTF8Encoding $false))
    if (Test-Path -LiteralPath $Path) {
        # PowerShell converts $null to "" for a string argument; File.Replace
        # rejects "" as a backup path. [NullString] passes a real null.
        [System.IO.File]::Replace($tmp, $Path, [NullString]::Value)
    }
    else {
        [System.IO.File]::Move($tmp, $Path)
    }
}

function Read-RingRun {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "Run file not found: $Path" }
    $run = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($null -eq $run -or $null -eq $run.PSObject.Properties['RunId'] -or $null -eq $run.PSObject.Properties['Rings']) {
        throw "Run file is not a ring run: $Path"
    }
    $run.Rings = @($run.Rings)
    return $run
}

function Get-RingRunRing {
    param([Parameter(Mandatory)]$Run, [Parameter(Mandatory)][int]$RingIndex)
    foreach ($r in @($Run.Rings)) { if ([int]$r.Index -eq $RingIndex) { return $r } }
    return $null
}

function Set-RingRunRingStatus {
    <#
    .SYNOPSIS
        Applies one status transition: Held -> Creating -> Created, a failed
        Creating -> Held retry, or Created -> Removed. Held -> Created remains
        accepted for compatibility with existing callers and run files.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates an in-memory run object only.')]
    param(
        [Parameter(Mandatory)]$Run,
        [Parameter(Mandatory)][int]$RingIndex,
        [Parameter(Mandatory)][ValidateSet('Creating', 'Held', 'Created', 'Removed')][string]$Status,
        [string]$DeploymentID,
        $AssignmentID,
        [datetime]$At = (Get-Date)
    )

    $ring = Get-RingRunRing -Run $Run -RingIndex $RingIndex
    if ($null -eq $ring) { throw "Ring $RingIndex is not in this run." }
    $current = [string]$ring.Status

    if ($Status -eq 'Creating') {
        if ($current -ne 'Held') { throw "Ring $RingIndex is $current; only a Held ring can become Creating." }
        $ring.Status       = 'Creating'
        $ring.CreatingAt   = ConvertTo-RingDateText $At
        $ring.DeploymentID = $null
        $ring.AssignmentID = $null
    }
    elseif ($Status -eq 'Held') {
        if ($current -ne 'Creating') { throw "Ring $RingIndex is $current; only a Creating ring can return to Held." }
        $ring.Status       = 'Held'
        $ring.CreatingAt   = $null
        $ring.DeploymentID = $null
        $ring.AssignmentID = $null
    }
    elseif ($Status -eq 'Created') {
        if ($current -notin @('Held', 'Creating')) { throw "Ring $RingIndex is $current; only a Held ring can become Created." }
        if ([string]::IsNullOrWhiteSpace($DeploymentID)) { throw "Ring $RingIndex needs a deployment ID to become Created." }
        $ring.Status       = 'Created'
        $ring.DeploymentID = $DeploymentID
        $ring.AssignmentID = if ($null -ne $AssignmentID) { [string]$AssignmentID } else { $null }
        $ring.CreatedAt    = ConvertTo-RingDateText $At
        $ring.CreatingAt   = $null
    }
    else {
        if ($current -ne 'Created') { throw "Ring $RingIndex is $current; only a Created ring can become Removed." }
        $ring.Status    = 'Removed'
        $ring.RemovedAt = ConvertTo-RingDateText $At
    }
    return $Run
}

function Get-RingRunNextHeld {
    <#
    .SYNOPSIS
        Returns the index of the first Held or Creating ring, or 0 when none
        is pending. Creating rings block later rings until reconciliation.
    #>
    param([Parameter(Mandatory)]$Run)
    foreach ($r in @($Run.Rings | Sort-Object { [int]$_.Index })) {
        if ([string]$r.Status -in @('Held', 'Creating')) { return [int]$r.Index }
    }
    return 0
}

function Test-RingRunFinished {
    <#
    .SYNOPSIS
        A run is finished when no ring is Held or Creating, or when the ring
        before the next pending ring was Removed (nothing can be promoted).
    #>
    param([Parameter(Mandatory)]$Run)
    $next = Get-RingRunNextHeld -Run $Run
    if ($next -eq 0) { return $true }
    $nextRing = Get-RingRunRing -Run $Run -RingIndex $next
    if ([string]$nextRing.Status -eq 'Creating') { return $false }
    $prev = Get-RingRunRing -Run $Run -RingIndex ($next - 1)
    return ($null -ne $prev -and [string]$prev.Status -eq 'Removed')
}

function Test-RingRunPromotable {
    <#
    .SYNOPSIS
        Decides whether a ring can be promoted, from a run object read fresh
        from disk. Returns @{ Ok; Reason }.
    #>
    param(
        [Parameter(Mandatory)]$Run,
        [Parameter(Mandatory)][int]$RingIndex
    )

    if ($Run.ClosedAt) { return @{ Ok = $false; Reason = 'The run is closed.' } }
    $ring = Get-RingRunRing -Run $Run -RingIndex $RingIndex
    if ($null -eq $ring) { return @{ Ok = $false; Reason = "Ring $RingIndex is not in this run." } }
    if ([string]$ring.Status -eq 'Creating') {
        return @{ Ok = $false; Reason = ("Ring {0} has an unresolved create attempt. Refresh and reconcile before trying again." -f $RingIndex) }
    }
    if ([string]$ring.Status -ne 'Held') {
        return @{ Ok = $false; Reason = ("Ring {0} is {1}, not Held. The run file changed after it was loaded." -f $RingIndex, $ring.Status) }
    }
    $next = Get-RingRunNextHeld -Run $Run
    if ($next -ne $RingIndex) {
        return @{ Ok = $false; Reason = ("Ring {0} is not the next held ring; ring {1} is." -f $RingIndex, $next) }
    }
    if ($RingIndex -gt 1) {
        $prev = Get-RingRunRing -Run $Run -RingIndex ($RingIndex - 1)
        if ([string]$prev.Status -ne 'Created') {
            return @{ Ok = $false; Reason = ("Ring {0} is {1}; Promote needs it Created." -f ($RingIndex - 1), $prev.Status) }
        }
    }
    return @{ Ok = $true; Reason = '' }
}

function Get-RingPromoteShift {
    <#
    .SYNOPSIS
        Returns the minutes to add to a held ring whose deadline passed, so it
        becomes available at the next whole minute and keeps its
        available-to-deadline gap. Returns 0 when the deadline is still ahead
        or there is none.
    #>
    param(
        [Parameter(Mandatory)]$Ring,
        [datetime]$Now = (Get-Date)
    )
    $deadline  = ConvertFrom-RingDateText $Ring.DeadlineDateTime
    $available = ConvertFrom-RingDateText $Ring.AvailableDateTime
    if ($null -eq $deadline -or $deadline -gt $Now) { return 0 }
    $start = $Now.Date.AddHours($Now.Hour).AddMinutes($Now.Minute + 1)
    return [int][math]::Ceiling(($start - $available).TotalMinutes)
}

function Move-RingRunSchedule {
    <#
    .SYNOPSIS
        Adds minutes to the dates of every Held ring from FromIndex on.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates an in-memory run object only.')]
    param(
        [Parameter(Mandatory)]$Run,
        [Parameter(Mandatory)][int]$FromIndex,
        [Parameter(Mandatory)][int]$Minutes
    )
    if ($Minutes -eq 0) { return $Run }
    foreach ($r in @($Run.Rings)) {
        if ([int]$r.Index -lt $FromIndex -or [string]$r.Status -ne 'Held') { continue }
        $r.AvailableDateTime = ConvertTo-RingDateText ((ConvertFrom-RingDateText $r.AvailableDateTime).AddMinutes($Minutes))
        if ($r.DeadlineDateTime) {
            $r.DeadlineDateTime = ConvertTo-RingDateText ((ConvertFrom-RingDateText $r.DeadlineDateTime).AddMinutes($Minutes))
        }
        $r.ShiftedMinutes = [int]$r.ShiftedMinutes + $Minutes
    }
    return $Run
}

function Enter-RingRunLock {
    <#
    .SYNOPSIS
        Takes <run>.lock with FileMode.CreateNew, which fails when the lock
        exists (also on an SMB share). Returns the lock path; throws when
        another session holds it.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $lock = $Path + '.lock'
    try {
        $fs = [System.IO.File]::Open($lock, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    }
    catch {
        $owner = 'an unknown session'
        try { $owner = [System.IO.File]::ReadAllText($lock) } catch { $null = $_ }
        throw ("The run file is locked by {0}. If nobody is working on this run, delete {1}." -f $owner, $lock)
    }
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes(("{0}\{1} on {2} at {3}" -f $env:USERDOMAIN, $env:USERNAME, $env:COMPUTERNAME, (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
        $fs.Write($bytes, 0, $bytes.Length)
    }
    finally { $fs.Dispose() }
    return $lock
}

function Exit-RingRunLock {
    param([AllowEmptyString()][AllowNull()][string]$LockPath)
    if ($LockPath -and (Test-Path -LiteralPath $LockPath)) { Remove-Item -LiteralPath $LockPath -Force -ErrorAction SilentlyContinue }
}

function Get-RingDeploymentSummary {
    <#
    .SYNOPSIS
        Reads live counts for one deployment with Get-CMDeployment
        -DeploymentId. Found is $true, $false (the site has no such
        deployment), or $null (the read failed; Error says why).
    #>
    param([AllowEmptyString()][AllowNull()][string]$DeploymentId)

    $empty = [ordered]@{ Found = $null; Summarized = $false; Targeted = 0; Success = 0; Errors = 0; InProgress = 0; Unknown = 0; Other = 0; SummarizationTime = $null; Error = $null }
    if ([string]::IsNullOrWhiteSpace($DeploymentId)) {
        $empty.Error = 'No deployment ID is recorded for this ring.'
        return [PSCustomObject]$empty
    }
    try {
        $d = @(Get-CMDeployment -DeploymentId $DeploymentId -ErrorAction Stop)
    }
    catch {
        $empty.Error = $_.Exception.Message
        Write-Log ("Deployment summary read failed for {0}: {1}" -f $DeploymentId, $_.Exception.Message) -Level WARN
        return [PSCustomObject]$empty
    }
    if ($d.Count -eq 0 -or $null -eq $d[0]) {
        $empty.Found = $false
        return [PSCustomObject]$empty
    }
    return ConvertTo-RingSummary -Deployment $d[0]
}

function ConvertTo-RingSummary {
    param([Parameter(Mandatory)]$Deployment, [string]$CorrectedDeploymentID)
    return [PSCustomObject][ordered]@{
        Found                 = $true
        Summarized            = $true
        Targeted              = [int]$Deployment.NumberTargeted
        Success               = [int]$Deployment.NumberSuccess
        Errors                = [int]$Deployment.NumberErrors
        InProgress            = [int]$Deployment.NumberInProgress
        Unknown               = [int]$Deployment.NumberUnknown
        Other                 = [int]$Deployment.NumberOther
        SummarizationTime     = $Deployment.SummarizationTime
        Error                 = $null
        CorrectedDeploymentID = if ($CorrectedDeploymentID) { $CorrectedDeploymentID } else { $null }
    }
}

function Test-RingDeploymentExist {
    <#
    .SYNOPSIS
        Reads the deployment object itself (SMS_ApplicationAssignment,
        SMS_UpdateGroupAssignment, or SMS_Advertisement) by its deployment ID,
        not its summary row. Returns @{ Exists; Error }: Exists is $true,
        $false, or $null when the read failed.
    #>
    param(
        [Parameter(Mandatory)][string]$DeploymentId,
        [Parameter(Mandatory)][ValidateSet('Application', 'Package', 'TaskSequence', 'SUG')][string]$ObjectType
    )

    try {
        $found = switch ($ObjectType) {
            'Application'  { Get-CMApplicationDeployment -DeploymentId $DeploymentId -ErrorAction Stop }
            'Package'      { Get-CMPackageDeployment -DeploymentId $DeploymentId -ErrorAction Stop }
            'TaskSequence' { Get-CMTaskSequenceDeployment -DeploymentId $DeploymentId -Fast -ErrorAction Stop }
            'SUG'          { Get-CMUpdateGroupDeployment -DeploymentId $DeploymentId -ErrorAction Stop }
        }
    }
    catch {
        return @{ Exists = $null; Error = $_.Exception.Message }
    }
    return @{ Exists = (@($found | Where-Object { $null -ne $_ }).Count -gt 0); Error = $null }
}

function Test-RingDeploymentIdForm {
    <#
    .SYNOPSIS
        True when the ID has the form the per-type cmdlets take: a GUID
        (AssignmentUniqueID) for an application or update group, an
        8-character AdvertisementID for a package or task sequence.
    #>
    param(
        [AllowEmptyString()][AllowNull()][string]$DeploymentId,
        [Parameter(Mandatory)][ValidateSet('Application', 'Package', 'TaskSequence', 'SUG')][string]$ObjectType
    )
    $id = ([string]$DeploymentId).Trim()
    if ($ObjectType -in @('Application', 'SUG')) {
        $g = [guid]::Empty
        return [guid]::TryParse($id, [ref]$g)
    }
    return ($id -match '^[A-Za-z0-9]{8}$')
}

function Get-RingLiveSummary {
    <#
    .SYNOPSIS
        Live counts for one created ring, in three reads:
          1. the summary row by deployment ID;
          2. the summary rows of the ring's collection, matched on the
             recorded DeploymentID or AssignmentID (a match under another
             DeploymentID returns CorrectedDeploymentID);
          3. the deployment object itself by deployment ID.
        The summary rows come from site summarization and can lag a new
        deployment, so only read 3 decides a deletion: Found = $false only
        when the object is gone. An object without a summary row returns
        Found = $true, Summarized = $false.
    #>
    param(
        [Parameter(Mandatory)]$Ring,
        [Parameter(Mandatory)][ValidateSet('Application', 'Package', 'TaskSequence', 'SUG')][string]$ObjectType
    )

    $recordedId = [string]$Ring.DeploymentID
    $s = Get-RingDeploymentSummary -DeploymentId $recordedId
    if ($s.Found -ne $false) { return $s }

    $collectionName = [string]$Ring.CollectionName
    if (-not [string]::IsNullOrWhiteSpace($collectionName)) {
        $feature = switch ($ObjectType) {
            'Application'  { 'Application' }
            'Package'      { 'Package' }
            'TaskSequence' { 'TaskSequence' }
            'SUG'          { 'SoftwareUpdate' }
        }
        try {
            $all = @(Get-CMDeployment -CollectionName $collectionName -FeatureType $feature -ErrorAction Stop)
        }
        catch {
            $s.Found = $null
            $s.Error = $_.Exception.Message
            return $s
        }
        $recordedAssignment = [string]$Ring.AssignmentID
        foreach ($d in $all) {
            if ($null -eq $d) { continue }
            $idMatch = ([string]$d.DeploymentID) -eq $recordedId
            $assignmentMatch = $recordedAssignment -and $recordedAssignment -ne '0' -and ([string]$d.AssignmentID) -eq $recordedAssignment
            if ($idMatch -or $assignmentMatch) {
                $corrected = if ($idMatch) { '' } else { [string]$d.DeploymentID }
                return ConvertTo-RingSummary -Deployment $d -CorrectedDeploymentID $corrected
            }
        }
    }

    if (-not (Test-RingDeploymentIdForm -DeploymentId $recordedId -ObjectType $ObjectType)) {
        $s.Found = $null
        $s.Error = ("Deployment ID '{0}' does not have the form of a {1} deployment ID, so a deletion cannot be confirmed." -f $recordedId, $ObjectType)
        return $s
    }
    $exists = Test-RingDeploymentExist -DeploymentId $recordedId -ObjectType $ObjectType
    if ($null -eq $exists.Exists) {
        $s.Found = $null
        $s.Error = $exists.Error
        return $s
    }
    if ($exists.Exists) {
        $s.Found = $true
        $s.Summarized = $false
        return $s
    }
    $s.Found = $false
    return $s
}

function Update-RingRunReconcile {
    <#
    .SYNOPSIS
        Reconciles Creating rings from live duplicate-check results, marks each
        Created ring Removed when its summary says the site has no deployment,
        and corrects IDs when needed. Failed reads never change a status.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates an in-memory run object only.')]
    param(
        [Parameter(Mandatory)]$Run,
        [Parameter(Mandatory)][hashtable]$Summaries,
        [hashtable]$CreatingCandidates = @{},
        [ValidateRange(0, 1440)][int]$RecoveryDelayMinutes = 5,
        [datetime]$At = (Get-Date)
    )

    $removed   = @()
    $corrected = @()
    $recovered = @()
    $reset     = @()
    $recoveryNotes = @{}

    foreach ($r in @($Run.Rings | Where-Object { [string]$_.Status -eq 'Creating' })) {
        $idx = [int]$r.Index
        if (-not $CreatingCandidates.ContainsKey($idx)) {
            $recoveryNotes[$idx] = 'Create outcome is unresolved. Refresh with a live site connection to reconcile it.'
            continue
        }
        $candidateCheck = $CreatingCandidates[$idx]
        if ($candidateCheck.CheckFailed) {
            $recoveryNotes[$idx] = ('Could not reconcile the create attempt: {0}' -f $candidateCheck.Error)
            continue
        }
        $found = @($candidateCheck.Deployments | Where-Object { $null -ne $_ })
        if ($found.Count -eq 1) {
            try {
                $reference = Get-RingDeploymentReference -Type ([string]$Run.Object.Type) -Deployment $found[0]
                [void](Set-RingRunRingStatus -Run $Run -RingIndex $idx -Status Created `
                    -DeploymentID $reference.DeploymentID -AssignmentID $reference.AssignmentID -At $At)
                $recovered += $idx
                $recoveryNotes[$idx] = ('Recovered the existing deployment {0} into the run.' -f $reference.DeploymentID)
            }
            catch {
                $recoveryNotes[$idx] = $_.Exception.Message
            }
            continue
        }
        if ($found.Count -gt 1) {
            $recoveryNotes[$idx] = ('Found {0} matching deployments. Remove extras, then refresh to recover this ring.' -f $found.Count)
            continue
        }

        $creatingAt = ConvertFrom-RingDateText $r.CreatingAt
        if ($null -ne $creatingAt -and ($At - $creatingAt).TotalMinutes -ge $RecoveryDelayMinutes) {
            [void](Set-RingRunRingStatus -Run $Run -RingIndex $idx -Status Held -At $At)
            $reset += $idx
            $recoveryNotes[$idx] = 'No matching deployment was found; the ring returned to Held and can be retried.'
        }
        else {
            $recoveryNotes[$idx] = 'No matching deployment is visible yet. Refresh again after the recovery wait before retrying.'
        }
    }

    foreach ($r in @($Run.Rings)) {
        if ([string]$r.Status -ne 'Created') { continue }
        $idx = [int]$r.Index
        if (-not $Summaries.ContainsKey($idx)) { continue }
        $s = $Summaries[$idx]
        if ($null -eq $s) { continue }
        if ($s.Found -eq $false) {
            [void](Set-RingRunRingStatus -Run $Run -RingIndex $idx -Status Removed -At $At)
            $removed += $idx
        }
        elseif ($s.Found -eq $true -and $s.PSObject.Properties['CorrectedDeploymentID'] -and $s.CorrectedDeploymentID) {
            $r.DeploymentID = [string]$s.CorrectedDeploymentID
            $corrected += $idx
        }
    }
    return @{
        Changed        = (($removed.Count + $corrected.Count + $recovered.Count + $reset.Count) -gt 0)
        Removed        = $removed
        Corrected      = $corrected
        Recovered      = $recovered
        Reset          = $reset
        RecoveryNotes  = $recoveryNotes
        Finished       = (Test-RingRunFinished -Run $Run)
    }
}

function Test-RingThreshold {
    <#
    .SYNOPSIS
        Compares a ring's live success rate with its threshold. An empty
        threshold is always met. Returns @{ Met; Percent; Reason }.
    #>
    param(
        [AllowNull()]$Summary,
        [AllowNull()]$ThresholdPercent
    )

    if ($null -eq $ThresholdPercent -or ([string]$ThresholdPercent).Trim().Length -eq 0) {
        return @{ Met = $true; Percent = $null; Reason = '' }
    }
    $t = [double]$ThresholdPercent
    if ($null -eq $Summary -or $Summary.Found -ne $true) {
        return @{ Met = $false; Percent = $null; Reason = ('No live counts; the {0}% threshold cannot be checked.' -f $t) }
    }
    if ($Summary.PSObject.Properties['Summarized'] -and $Summary.Summarized -eq $false) {
        return @{ Met = $false; Percent = $null; Reason = ('The site has not summarized this ring yet; the {0}% threshold cannot be checked.' -f $t) }
    }
    if ([int]$Summary.Targeted -le 0) {
        return @{ Met = $false; Percent = 0; Reason = ('No clients targeted yet; the threshold is {0}%.' -f $t) }
    }
    $pct = [math]::Round(([double]$Summary.Success * 100.0) / [double]$Summary.Targeted, 1)
    if ($pct -ge $t) { return @{ Met = $true; Percent = $pct; Reason = '' } }
    return @{ Met = $false; Percent = $pct; Reason = ('Success {0}% is below the {1}% threshold.' -f $pct, $t) }
}

function Get-RingRunFile {
    <#
    .SYNOPSIS
        Lists run files: open runs sit in the folder, closed runs in its
        closed subfolder.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Closed
    )
    $dir = if ($Closed) { Join-Path $Path 'closed' } else { $Path }
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    return @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)
}

function Close-RingRun {
    <#
    .SYNOPSIS
        Stamps ClosedAt, saves, and moves the run file to the closed
        subfolder. Returns the new path.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Moves the tool''s own run-state file.')]
    param(
        [Parameter(Mandatory)]$Run,
        [Parameter(Mandatory)][string]$Path,
        [datetime]$At = (Get-Date)
    )

    $Run.ClosedAt = ConvertTo-RingDateText $At
    Save-RingRun -Run $Run -Path $Path
    $closedDir = Join-Path (Split-Path -Path $Path -Parent) 'closed'
    if (-not (Test-Path -LiteralPath $closedDir)) { New-Item -ItemType Directory -Path $closedDir -Force | Out-Null }
    $dest = Get-RingFreePath -Path (Join-Path $closedDir (Split-Path -Path $Path -Leaf))
    # Move without -Force: an existing closed file is never replaced.
    Move-Item -LiteralPath $Path -Destination $dest
    Write-Log "Ring run closed: $dest"
    return $dest
}

# ---------------------------------------------------------------------------
# Ring deployment (one call to the per-type deployment function)
# ---------------------------------------------------------------------------

function Invoke-RingDeployment {
    <#
    .SYNOPSIS
        Creates one ring's deployment through the existing per-type function
        with the ring's computed dates and options. Returns that function's
        result hashtable.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('Application', 'Package', 'TaskSequence', 'SUG')][string]$Type,
        [Parameter(Mandatory)]$TargetObject,
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Ring,
        [string]$ProgramName,
        [ValidateSet('LocalTime', 'Utc')][string]$TimeBasedOn = 'LocalTime'
    )

    $available = ConvertFrom-RingDateText $Ring.AvailableDateTime
    $deadline  = ConvertFrom-RingDateText $Ring.DeadlineDateTime
    $purpose   = [string]$Ring.Purpose

    switch ($Type) {
        'Application' {
            $p = @{
                Application                = $TargetObject
                Collection                 = $Collection
                DeployPurpose              = $purpose
                AvailableDateTime          = $available
                TimeBasedOn                = $TimeBasedOn
                UserNotification           = [string]$Ring.UserNotification
                OverrideServiceWindow      = [bool]$Ring.OverrideServiceWindow
                RebootOutsideServiceWindow = [bool]$Ring.RebootOutsideServiceWindow
                UseMeteredNetwork          = [bool]$Ring.AllowMeteredConnection
            }
            if ($deadline) { $p['DeadlineDateTime'] = $deadline }
            return Invoke-ApplicationDeployment @p
        }
        'Package' {
            $p = @{
                Package                    = $TargetObject
                ProgramName                = $ProgramName
                Collection                 = $Collection
                DeployPurpose              = $purpose
                AvailableDateTime          = $available
                TimeBasedOn                = $TimeBasedOn
                OverrideServiceWindow      = [bool]$Ring.OverrideServiceWindow
                RebootOutsideServiceWindow = [bool]$Ring.RebootOutsideServiceWindow
                UseMeteredNetwork          = [bool]$Ring.AllowMeteredConnection
                FastNetworkOption          = [string]$Ring.FastNetworkOption
                SlowNetworkOption          = [string]$Ring.SlowNetworkOption
                RerunBehavior              = [string]$Ring.RerunBehavior
            }
            if ($deadline) { $p['DeadlineDateTime'] = $deadline }
            return Invoke-PackageDeployment @p
        }
        'TaskSequence' {
            $p = @{
                TaskSequence               = $TargetObject
                Collection                 = $Collection
                DeployPurpose              = $purpose
                AvailableDateTime          = $available
                Availability               = [string]$Ring.TaskSequenceAvailability
                TimeBasedOn                = $TimeBasedOn
                ShowTaskSequenceProgress   = [bool]$Ring.ShowTaskSequenceProgress
                OverrideServiceWindow      = [bool]$Ring.OverrideServiceWindow
                RebootOutsideServiceWindow = [bool]$Ring.RebootOutsideServiceWindow
                UseMeteredNetwork          = [bool]$Ring.AllowMeteredConnection
            }
            if ($deadline) { $p['DeadlineDateTime'] = $deadline }
            return Invoke-TaskSequenceDeployment @p
        }
        'SUG' {
            $p = @{
                SUG                         = $TargetObject
                Collection                  = $Collection
                DeployPurpose               = $purpose
                AvailableDateTime           = $available
                TimeBasedOn                 = $TimeBasedOn
                UserNotification            = [string]$Ring.UserNotification
                SoftwareInstallation        = [bool]$Ring.OverrideServiceWindow
                AllowRestart                = [bool]$Ring.RebootOutsideServiceWindow
                UseMeteredNetwork           = [bool]$Ring.AllowMeteredConnection
                AllowBoundaryFallback       = [bool]$Ring.AllowBoundaryFallback
                DownloadFromMicrosoftUpdate = [bool]$Ring.AllowMicrosoftUpdate
                RequirePostRebootFullScan   = [bool]$Ring.RequirePostRebootFullScan
            }
            if ($deadline) { $p['DeadlineDateTime'] = $deadline }
            return Invoke-SUGDeployment @p
        }
    }
}

function New-RingAuditRecord {
    <#
    .SYNOPSIS
        Builds the Write-DeploymentLog record for one ring: the fields the
        single-deployment flow writes, plus PlanName, RingIndex, RingName,
        and RunId.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Builds an in-memory audit record.')]
    param(
        [Parameter(Mandatory)][ValidateSet('Application', 'Package', 'TaskSequence', 'SUG')][string]$Type,
        [Parameter(Mandatory)]$TargetObject,
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)]$Ring,
        [Parameter(Mandatory)][hashtable]$Result,
        [Parameter(Mandatory)][string]$PlanName,
        [Parameter(Mandatory)][string]$RunId,
        [string]$ProgramName
    )

    switch ($Type) {
        'Application'  { $name = $TargetObject.LocalizedDisplayName; $version = $TargetObject.SoftwareVersion }
        'Package'      { $name = $TargetObject.Name;                 $version = $ProgramName }
        'TaskSequence' { $name = $TargetObject.Name;                 $version = [string]$Ring.TaskSequenceAvailability }
        'SUG'          { $name = $TargetObject.LocalizedDisplayName; $version = "($($TargetObject.NumberOfUpdates) updates)" }
    }
    $deadline = ConvertFrom-RingDateText $Ring.DeadlineDateTime
    return @{
        DeploymentType     = $Type
        ApplicationName    = $name
        ApplicationVersion = $version
        CollectionName     = $Collection.Name
        CollectionID       = $Collection.CollectionID
        MemberCount        = $Collection.MemberCount
        DeployPurpose      = [string]$Ring.Purpose
        DeadlineDateTime   = if ($deadline) { $deadline.ToString('yyyy-MM-ddTHH:mm:ss') } else { $null }
        DeploymentID       = $Result.DeploymentID
        Result             = if ($Result.Success) { 'Success' } else { ('Failed: {0}' -f $Result.Error) }
        PlanName           = $PlanName
        RingIndex          = [int]$Ring.Index
        RingName           = [string]$Ring.Name
        RunId              = $RunId
    }
}

function Test-RingPreflight {
    <#
    .SYNOPSIS
        Runs the five pre-execution checks for each ring: object exists,
        content distributed, collection valid (by CollectionID), collection
        safe, no duplicate deployment. Object and content checks run once and
        count for every ring. Returns @{ Ok; Object; ObjectMessage; Rings }
        where each ring entry is @{ Index; Collection; CollectionSafe; Passed;
        Ok; Message; DuplicateCheckError; ExistingDeployments }.

        -SkipContentCheck counts the content check as passed without a read.
        Rechecks at create time use it: distribution to a ring's new DP group
        leaves the content in progress, and a content check there would block
        that ring and every later ring of the run.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('Application', 'Package', 'TaskSequence', 'SUG')][string]$Type,
        [Parameter(Mandatory)][string]$ObjectName,
        [Parameter(Mandatory)][AllowEmptyCollection()][array]$Rings,
        [string]$ProgramName,
        [string]$ExpectedObjectId,
        [switch]$SkipContentCheck
    )

    $object = switch ($Type) {
        'Application'  { Test-ApplicationExists -ApplicationName $ObjectName }
        'Package'      { Test-PackageExists -PackageName $ObjectName }
        'TaskSequence' { Test-TaskSequenceExists -TaskSequenceName $ObjectName }
        'SUG'          { Test-SUGExists -SUGName $ObjectName }
    }

    $objectOk = $null -ne $object
    $objectMessage = if ($objectOk) { '' } else { ("{0} '{1}' was not found." -f $Type, $ObjectName) }

    if ($objectOk -and $ExpectedObjectId) {
        $identity = Get-RingObjectIdentity -Type $Type -TargetObject $object -ProgramName $ProgramName
        if ($identity.ID -ne $ExpectedObjectId) {
            $objectOk = $false
            $objectMessage = ("{0} '{1}' now has ID {2}; the run was created for {3}." -f $Type, $ObjectName, $identity.ID, $ExpectedObjectId)
        }
    }
    if ($objectOk -and $Type -eq 'Package') {
        $programs = @(Get-CMPackagePrograms -Package $object | ForEach-Object { [string]$_.ProgramName })
        if ([string]::IsNullOrWhiteSpace($ProgramName) -or $ProgramName -notin $programs) {
            $objectOk = $false
            $objectMessage = ("Package '{0}' has no program '{1}'." -f $ObjectName, $ProgramName)
        }
    }

    $contentOk = $true
    $contentMessage = ''
    if ($objectOk -and -not $SkipContentCheck -and $Type -in @('Application', 'Package')) {
        $dist = Test-ContentDistributed -Application $object
        if (-not $dist.IsFullyDistributed) {
            $contentOk = $false
            $contentMessage = ('Content is not fully distributed: {0}/{1} DP(s) succeeded.' -f $dist.NumberSuccess, $dist.Targeted)
        }
    }

    $ringResults = foreach ($r in @($Rings)) {
        $passed = 0
        $message = ''
        $collection = $null
        $collectionSafe = $false
        $duplicateCheckError = ''
        $existingDeployments = @()
        if ($objectOk) { $passed++ } else { $message = $objectMessage }
        if ($contentOk) { $passed++ } elseif (-not $message) { $message = $contentMessage }

        $cid = ([string]$r.CollectionID).Trim().ToUpperInvariant()
        if ($cid) { $collection = Test-CollectionValid -CollectionId $cid }
        if ($null -ne $collection) {
            $passed++
            $safe = Test-CollectionSafe -Collection $collection
            if ($safe.IsSafe) {
                $collectionSafe = $true
                $passed++
                $dup = if ($objectOk) {
                    switch ($Type) {
                        'Application'  { Test-DuplicateDeployment -ApplicationName $object.LocalizedDisplayName -CollectionName $collection.Name }
                        'Package'      { Test-DuplicatePackageDeployment -PackageID $object.PackageID -ProgramName $ProgramName -CollectionName $collection.Name }
                        'TaskSequence' { Test-DuplicateTaskSequenceDeployment -TaskSequencePackageId $object.PackageID -CollectionName $collection.Name }
                        'SUG'          { Test-DuplicateSUGDeployment -SUGName $object.LocalizedDisplayName -CollectionName $collection.Name }
                    }
                } else { $null }
                $dupFailure = Get-DuplicateCheckFailure -Result $dup
                if ($objectOk -and $null -ne $dupFailure) {
                    $duplicateCheckError = [string]$dupFailure.Error
                    if (-not $message) { $message = ("Could not check for an existing deployment: {0}" -f $duplicateCheckError) }
                }
                elseif ($objectOk -and $null -eq $dup) { $passed++ }
                elseif ($objectOk) {
                    $existingDeployments = @($dup)
                    if (-not $message) { $message = ("A deployment of this object to '{0}' already exists." -f $collection.Name) }
                }
            }
            elseif (-not $message) { $message = $safe.Reason }
        }
        elseif (-not $message) {
            $message = if ($cid) { ("Collection {0} was not found or is not a device collection." -f $cid) } else { 'No target collection.' }
        }

        @{
            Index      = [int]$r.Index
            Collection = $collection
            CollectionSafe = $collectionSafe
            Passed     = $passed
            Ok         = ($passed -eq 5)
            Message    = if ($passed -eq 5) { '5/5 checks passed' } else { ('{0}/5: {1}' -f $passed, $message) }
            DuplicateCheckError = $duplicateCheckError
            ExistingDeployments = $existingDeployments
        }
    }

    $ringResults = @($ringResults)
    return @{
        Ok            = ($objectOk -and $contentOk -and $ringResults.Count -gt 0 -and @($ringResults | Where-Object { -not $_.Ok }).Count -eq 0)
        Object        = $object
        ObjectMessage = if ($objectMessage) { $objectMessage } else { $contentMessage }
        Rings         = $ringResults
    }
}
