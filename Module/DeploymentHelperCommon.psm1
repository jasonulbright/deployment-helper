<#
.SYNOPSIS
    Core module for Deployment Helper.

.DESCRIPTION
    Import this module to get:
      - Structured logging (Initialize-Logging, Write-Log)
      - CM site connection management (Connect-CMSite, Disconnect-CMSite, Test-CMConnection)
      - Pre-execution validation (Test-ApplicationExists, Test-ContentDistributed, Test-CollectionValid, Test-CollectionSafe, Test-DuplicateDeployment)
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
        $existing = Get-CMDeployment -SoftwareName $ApplicationName -CollectionName $CollectionName -ErrorAction Stop
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
        [ValidateSet('DisplayAll','DisplaySoftwareCenterOnly','HideAll')]
        [string]$UserNotification = 'DisplayAll',
        [bool]$OverrideServiceWindow = $false,
        [bool]$RebootOutsideServiceWindow = $false,
        [bool]$AllowMeteredConnection = $false,
        [string]$Comment = ''
    )

    $params = @{
        Name                        = $Application.LocalizedDisplayName
        CollectionName              = $Collection.Name
        DeployPurpose               = $DeployPurpose
        DeployAction                = 'Install'
        AvailableDateTime           = $AvailableDateTime
        TimeBaseOn                  = 'LocalTime'
        UserNotification            = $UserNotification
        OverrideServiceWindow       = $OverrideServiceWindow
        RebootOutsideServiceWindow  = $RebootOutsideServiceWindow
        ErrorAction                 = 'Stop'
    }

    if ($DeployPurpose -eq 'Required' -and $DeadlineDateTime) {
        $params['DeadlineDateTime'] = $DeadlineDateTime
    }
    if ($Comment) {
        $params['Comment'] = $Comment
    }
    if ($AllowMeteredConnection) {
        $params['AllowMeteredConnection'] = $true
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
        [ValidateSet('DisplayAll','DisplaySoftwareCenterOnly','HideAll')]
        [string]$UserNotification = 'DisplayAll',
        [bool]$SoftwareInstallation = $false,
        [bool]$AllowRestart = $false,
        [bool]$AllowUseMeteredNetwork = $false,
        [string]$Comment = ''
    )

    $params = @{
        SoftwareUpdateGroupName = $SUG.LocalizedDisplayName
        CollectionName          = $Collection.Name
        DeploymentType          = $DeployPurpose
        AvailableDateTime       = $AvailableDateTime
        TimeBasedOn             = 'LocalTime'
        UserNotification        = $UserNotification
        SoftwareInstallation    = $SoftwareInstallation
        AllowRestart            = $AllowRestart
        ErrorAction             = 'Stop'
    }

    if ($DeployPurpose -eq 'Required' -and $DeadlineDateTime) {
        $params['DeadlineDateTime'] = $DeadlineDateTime
    }
    if ($Comment) {
        $params['DeploymentName'] = $Comment
    }

    # Required SUG: force download fallback settings
    if ($DeployPurpose -eq 'Required') {
        $params['ProtectedType']    = 'RemoteDistributionPoint'
        $params['UnprotectedType']  = 'UnprotectedDistributionPoint'
    }

    if ($AllowUseMeteredNetwork) {
        $params['UseBranchCache'] = $true
        $params['AllowUseMeteredNetwork'] = $true
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
        DeployPurpose               = $Config.DeployPurpose
        UserNotification            = $Config.UserNotification
        OverrideServiceWindow       = $Config.OverrideServiceWindow
        RebootOutsideServiceWindow  = $Config.RebootOutsideServiceWindow
        AllowMeteredConnection      = $Config.AllowMeteredConnection
        DefaultDeadlineOffsetHours  = $Config.DefaultDeadlineOffsetHours
    }

    $template | ConvertTo-Json | Set-Content -LiteralPath $TemplatePath -Encoding UTF8
    Write-Log "Saved deployment template '$TemplateName' to $TemplatePath"
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
        ChangeTicket       = $Record.ChangeTicket
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
        Comment            = $Record.Comment
    }

    $json = $entry | ConvertTo-Json -Compress
    Add-Content -LiteralPath $LogPath -Value $json -Encoding UTF8
    Write-Log "Deployment log entry written to $LogPath"
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
        '  .success { color: #228B22; font-weight: bold; }',
        '  .failed { color: #B40000; font-weight: bold; }',
        '</style>'
    ) -join "`r`n"

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $headerHtml = "<h1>Deployment History Report</h1><div class='subtitle'>Generated: $timestamp</div>"

    $columns = @('Timestamp','User','ChangeTicket','ApplicationName','ApplicationVersion','CollectionName','MemberCount','DeployPurpose','DeadlineDateTime','DeploymentID','Result')
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
