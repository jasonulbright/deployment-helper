<#
.SYNOPSIS
    Core module for Deployment Helper.

.DESCRIPTION
    Import this module to get:
      - Structured logging (Initialize-Logging, Write-Log)
      - CM site connection management (Connect-CMSite, Disconnect-CMSite, Test-CMConnection)
      - Pre-execution validation (Test-ApplicationExists, Test-ContentDistributed, Test-CollectionValid, Test-CollectionSafe, Test-DuplicateDeployment)
      - DP group management (Get-DPGroupList, Start-ContentDistributionToGroups)
      - Deployment preview and execution (Get-DeploymentPreview, Invoke-ApplicationDeployment)
      - Immutable deployment audit log (Write-DeploymentLog, Get-DeploymentHistory)
      - Deployment templates (Get-DeploymentTemplates)
      - Export to CSV and HTML (Export-DeploymentHistoryCsv, Export-DeploymentHistoryHtml)

.EXAMPLE
    Import-Module "$PSScriptRoot\Module\DeploymentHelperCommon.psd1" -Force
    Initialize-Logging -LogPath "C:\temp\dh.log"
    Connect-CMSite -SiteCode 'MCM' -SMSProvider 'sccm.domain.com'
#>

# ---------------------------------------------------------------------------
# Module-scoped state
# ---------------------------------------------------------------------------

$script:__DHLogPath             = $null
$script:OriginalLocation        = $null
$script:ConnectedSiteCode       = $null
$script:ConnectedSMSProvider    = $null

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Initialize-Logging {
    param([string]$LogPath)

    $script:__DHLogPath = $LogPath

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, severity-tagged log message.

    .DESCRIPTION
        INFO  -> Write-Host (stdout)
        WARN  -> Write-Host (stdout)
        ERROR -> Write-Host (stdout) + $host.UI.WriteErrorLine (stderr)

        -Quiet suppresses all console output but still writes to the log file.
    #>
    param(
        [AllowEmptyString()]
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',

        [switch]$Quiet
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[{0}] [{1,-5}] {2}" -f $timestamp, $Level, $Message

    if (-not $Quiet) {
        Write-Host $formatted

        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__DHLogPath) {
        Add-Content -LiteralPath $script:__DHLogPath -Value $formatted -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# CM Connection
# ---------------------------------------------------------------------------

function Connect-CMSite {
    <#
    .SYNOPSIS
        Imports the ConfigurationManager module, creates a PSDrive, and sets location.

    .DESCRIPTION
        Saves original location for restoration via Disconnect-CMSite.
        Returns $true on success, $false on failure.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$SMSProvider
    )

    $script:OriginalLocation = Get-Location

    # Import CM module if not already loaded
    if (-not (Get-Module ConfigurationManager -ErrorAction SilentlyContinue)) {
        $cmModulePath = $null
        if ($env:SMS_ADMIN_UI_PATH) {
            $cmModulePath = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
        }

        if (-not $cmModulePath -or -not (Test-Path -LiteralPath $cmModulePath)) {
            Write-Log "ConfigurationManager module not found. Ensure the CM console is installed." -Level ERROR
            return $false
        }

        try {
            Import-Module $cmModulePath -ErrorAction Stop
            Write-Log "Imported ConfigurationManager module"
        }
        catch {
            Write-Log "Failed to import ConfigurationManager module: $_" -Level ERROR
            return $false
        }
    }

    # Create PSDrive if needed
    if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        try {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SMSProvider -ErrorAction Stop | Out-Null
            Write-Log "Created PSDrive for site $SiteCode"
        }
        catch {
            Write-Log "Failed to create PSDrive for site $SiteCode : $_" -Level ERROR
            return $false
        }
    }

    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
        $site = Get-CMSite -SiteCode $SiteCode -ErrorAction Stop
        Write-Log "Connected to site $SiteCode ($($site.SiteName))"
        $script:ConnectedSiteCode    = $SiteCode
        $script:ConnectedSMSProvider = $SMSProvider
        return $true
    }
    catch {
        Write-Log "Failed to connect to site $SiteCode : $_" -Level ERROR
        return $false
    }
}

function Disconnect-CMSite {
    <#
    .SYNOPSIS
        Restores the original location before CM connection.
    #>
    if ($script:OriginalLocation) {
        try { Set-Location $script:OriginalLocation -ErrorAction SilentlyContinue } catch { }
    }
    $script:ConnectedSiteCode    = $null
    $script:ConnectedSMSProvider = $null
    Write-Log "Disconnected from CM site"
}

function Test-CMConnection {
    <#
    .SYNOPSIS
        Returns $true if currently connected to a CM site.
    #>
    if (-not $script:ConnectedSiteCode) { return $false }

    try {
        $drive = Get-PSDrive -Name $script:ConnectedSiteCode -PSProvider CMSite -ErrorAction Stop
        return ($null -ne $drive)
    }
    catch {
        return $false
    }
}

# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

function Test-ApplicationExists {
    param([Parameter(Mandatory)][string]$ApplicationName)

    try {
        $app = Get-CMApplication -Name $ApplicationName -ErrorAction Stop
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
    param([Parameter(Mandatory)][string]$CollectionName)

    try {
        $col = Get-CMCollection -Name $CollectionName -ErrorAction Stop
        if ($null -eq $col) {
            Write-Log "Collection not found: $CollectionName" -Level WARN
            return $null
        }
        if ($col.CollectionType -ne 2) {
            Write-Log "Collection '$CollectionName' is a User collection, not Device. Deployment requires a Device collection." -Level WARN
            return $null
        }
        Write-Log "Collection found: $CollectionName (ID: $($col.CollectionID), Members: $($col.MemberCount))"
        return $col
    }
    catch {
        Write-Log "Error querying collection '$CollectionName': $_" -Level ERROR
        return $null
    }
}

function Test-CollectionSafe {
    param([Parameter(Mandatory)]$Collection)

    $collectionId = $Collection.CollectionID

    if ($collectionId -match '^SMS000') {
        Write-Log "BLOCKED: Collection '$($Collection.Name)' ($collectionId) is a built-in system collection. Deployment not allowed." -Level ERROR
        return @{ IsSafe = $false; Reason = "Built-in system collection ($collectionId) is blocked for safety." }
    }

    Write-Log "Collection '$($Collection.Name)' ($collectionId) passed safety check"
    return @{ IsSafe = $true; Reason = '' }
}

function Test-DuplicateDeployment {
    param(
        [Parameter(Mandatory)][string]$ApplicationName,
        [Parameter(Mandatory)][string]$CollectionName
    )

    try {
        $existing = Get-CMApplicationDeployment -Name $ApplicationName -CollectionName $CollectionName -ErrorAction Stop
        if ($null -ne $existing -and @($existing).Count -gt 0) {
            Write-Log "Duplicate deployment found: '$ApplicationName' already deployed to '$CollectionName'" -Level WARN
            return $existing
        }
        Write-Log "No duplicate deployment: '$ApplicationName' to '$CollectionName'"
        return $null
    }
    catch {
        Write-Log "Error checking for duplicate deployment: $_" -Level ERROR
        return $null
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
        $sug = Get-CMSoftwareUpdateGroup -Name $SUGName -ErrorAction Stop
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
            Success      = $true
            DeploymentID = $deployment.AssignmentID
            Error        = $null
        }
    }
    catch {
        Write-Log "Deployment FAILED: $_" -Level ERROR
        return @{
            Success      = $false
            DeploymentID = $null
            Error        = $_.ToString()
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
            Success      = $true
            DeploymentID = $deployment.AssignmentID
            Error        = $null
        }
    }
    catch {
        Write-Log "SUG deployment FAILED: $_" -Level ERROR
        return @{
            Success      = $false
            DeploymentID = $null
            Error        = $_.ToString()
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

function Write-DeploymentLog {
    param(
        [Parameter(Mandatory)][string]$LogPath,
        [Parameter(Mandatory)][hashtable]$Record
    )

    $parentDir = Split-Path -Path $LogPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
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

    $json = $entry | ConvertTo-Json -Compress
    # A locked JSONL (user has it open in an editor) or a read-only path
    # can make Add-Content throw AFTER the deployment already succeeded.
    # Swallow the append failure into a WARN so the crash doesn't make it
    # look like the deployment itself failed; the cmdlet call happened
    # before this point and the CM assignment ID is already in the GUI log.
    try {
        Add-Content -LiteralPath $LogPath -Value $json -Encoding UTF8 -ErrorAction Stop
        Write-Log "Deployment log entry written to $LogPath"
    }
    catch {
        Write-Log ("Audit log append failed ({0}): {1}" -f $LogPath, $_) -Level WARN
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

    $Records | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
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

    $columns = @('Timestamp','User','DeploymentType','ApplicationName','ApplicationVersion','CollectionName','MemberCount','DeployPurpose','DeadlineDateTime','DeploymentID','Result')
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
        $pkg = Get-CMPackage -Name $PackageName -Fast -ErrorAction Stop
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
        $existing = Get-CMPackageDeployment -PackageId $PackageID -ProgramName $ProgramName -CollectionName $CollectionName -ErrorAction Stop
        if ($null -ne $existing -and @($existing).Count -gt 0) {
            Write-Log "Duplicate package deployment: PackageID=$PackageID Program='$ProgramName' Collection='$CollectionName'" -Level WARN
            return $existing
        }
        Write-Log "No duplicate package deployment: PackageID=$PackageID Program='$ProgramName' Collection='$CollectionName'"
        return $null
    }
    catch {
        Write-Log "Error checking duplicate package deployment: $_" -Level ERROR
        return $null
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
        $existing = Get-CMSoftwareUpdateDeployment -Name $SUGName -CollectionName $CollectionName -ErrorAction Stop
        if ($null -ne $existing -and @($existing).Count -gt 0) {
            Write-Log "Duplicate SUG deployment: SUG='$SUGName' Collection='$CollectionName'" -Level WARN
            return $existing
        }
        Write-Log "No duplicate SUG deployment: SUG='$SUGName' Collection='$CollectionName'"
        return $null
    }
    catch {
        Write-Log "Error checking duplicate SUG deployment: $_" -Level ERROR
        return $null
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
        $ts = Get-CMTaskSequence -Name $TaskSequenceName -Fast -ErrorAction Stop
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
        $existing = Get-CMTaskSequenceDeployment -TaskSequenceId $TaskSequencePackageId -CollectionName $CollectionName -Fast -ErrorAction Stop
        if ($null -ne $existing -and @($existing).Count -gt 0) {
            Write-Log "Duplicate TS deployment: TaskSequencePackageId=$TaskSequencePackageId Collection='$CollectionName'" -Level WARN
            return $existing
        }
        Write-Log "No duplicate TS deployment: TaskSequencePackageId=$TaskSequencePackageId Collection='$CollectionName'"
        return $null
    }
    catch {
        Write-Log "Error checking duplicate TS deployment: $_" -Level ERROR
        return $null
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
    if ($DeployPurpose -eq 'Required' -and $DeadlineDateTime) {
        $params['DeadlineDateTime'] = $DeadlineDateTime
    }
    if ($OverrideServiceWindow)      { $params['SoftwareInstallation'] = $true }
    if ($RebootOutsideServiceWindow) { $params['SystemRestart']        = $true }
    # UseMeteredNetwork is a Required-only param. Passing it on Available
    # TS deploys triggers a cmdlet WARN that leaks into the audit log
    # (MECM surfaces "Parameter X does not apply to deployments with Purpose
    # Available"). Apps + Packages already model this gating via their
    # $params hashtable conditions; mirror it here.
    if ($DeployPurpose -eq 'Required' -and $UseMeteredNetwork) {
        $params['UseMeteredNetwork'] = $true
    }

    try {
        Write-Log ("Executing TS deployment: {0} -> {1} ({2} devices) as {3}, Availability={4}" -f
            $TaskSequence.Name, $Collection.Name, $Collection.MemberCount, $DeployPurpose, $Availability)

        $deployment = New-CMTaskSequenceDeployment @params

        # Same fix as packages: New-CMTaskSequenceDeployment returns
        # SMS_Advertisement -- AdvertisementID, not DeploymentID.
        Write-Log "TS deployment created (AdvertisementID: $($deployment.AdvertisementID))"
        return @{
            Success      = $true
            DeploymentID = $deployment.AdvertisementID
            Error        = $null
        }
    }
    catch {
        Write-Log "TS deployment FAILED: $_" -Level ERROR
        return @{
            Success      = $false
            DeploymentID = $null
            Error        = $_.ToString()
        }
    }
}

function Invoke-PackageDeployment {
    <#
    .SYNOPSIS
        Deploys a legacy package + program to a device collection using New-CMPackageDeployment.

    .DESCRIPTION
        Wraps New-CMPackageDeployment with the standard-program parameter set.
        Success/failure is returned as a hashtable with DeploymentID.

        UTC semantics: TimeBasedOn='Utc' sets both -UseUtcForAvailableSchedule and
        -UseUtcForExpireSchedule. The package cmdlet treats these independently in MECM;
        for safety-critical tool usage we keep them in lockstep.

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
    if ($DeployPurpose -eq 'Required' -and $DeadlineDateTime) {
        $params['DeadlineDateTime'] = $DeadlineDateTime
    }
    if ($OverrideServiceWindow)      { $params['SoftwareInstallation'] = $true }
    if ($RebootOutsideServiceWindow) { $params['SystemRestart']        = $true }
    if ($UseMeteredNetwork)          { $params['UseMeteredNetwork']    = $true }

    try {
        Write-Log ("Executing package deployment: {0} / {1} -> {2} ({3} devices) as {4}" -f
            $Package.Name, $ProgramName, $Collection.Name, $Collection.MemberCount, $DeployPurpose)

        $deployment = New-CMPackageDeployment @params

        # New-CMPackageDeployment returns SMS_Advertisement (legacy pkg
        # deployment shape). The ID property is AdvertisementID; the
        # earlier $deployment.DeploymentID read was returning $null.
        Write-Log "Package deployment created (AdvertisementID: $($deployment.AdvertisementID))"
        return @{
            Success      = $true
            DeploymentID = $deployment.AdvertisementID
            Error        = $null
        }
    }
    catch {
        Write-Log "Package deployment FAILED: $_" -Level ERROR
        return @{
            Success      = $false
            DeploymentID = $null
            Error        = $_.ToString()
        }
    }
}
