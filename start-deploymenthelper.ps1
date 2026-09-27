<#
.SYNOPSIS
    Main window of Deployment Helper, a tool that creates Configuration Manager deployments with validation and an audit log.

.DESCRIPTION
    MahApps.Metro 2.4.10 WPF shell for deploying Apps, Packages, Task Sequences,
    and Software Update Groups to ConfigMgr device collections.

    Four deployment types surface via sidebar buttons. All types share a unified
    form pane in the right column. Session 2 fully wires the Apps type
    (5-check validation plus Invoke-ApplicationDeployment). Packages, Task
    Sequences, and Software Update Groups ship their wiring in later sessions.

    The Ring deployment entry expands one object through a ring plan
    (Rings\*.json) into one ordinary deployment per ring.

    Shared pre-execution validation for Apps: Test-ApplicationExists,
    Test-ContentDistributed, Test-CollectionValid, Test-CollectionSafe,
    Test-DuplicateDeployment. All implemented in
    Module/DeploymentHelperCommon.psm1 (preserved as-is).

.PARAMETER SiteCode
    ConfigMgr site code (three alphanumeric characters). Optional on the
    command line; when omitted, the value from DeploymentHelper.prefs.json
    is used, and Options > Connection lets the user set it at runtime.

.PARAMETER SMSProvider
    Fully qualified SMS Provider host name. Optional on the command line;
    when omitted, the value from DeploymentHelper.prefs.json is used.

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.7.2+
      - MahApps.Metro 2.4.10 DLLs in .\Lib\
      - ConfigurationManager admin console (for CM cmdlets)

    ScriptName : start-deploymenthelper.ps1
    Version    : 2026.09.27.0011
    Updated    : 2026-09-27
#>

param(
    # Empty defaults: the shipped app should not assume any particular
    # lab or production site. On first launch the user sets these via
    # Options > Connection; the values persist to
    # DeploymentHelper.prefs.json next to the script.
    [string]$SiteCode    = '',
    [string]$SMSProvider = ''
)

$ErrorActionPreference = 'Stop'

# =============================================================================
# Crash capture -- if anything throws at startup the user otherwise sees
# "powershell crashed" with no log. Transcript + catch-all => we get a file
# to read. Path is per-launch so transcripts never collide.
# =============================================================================
try {
    $__tx = $null
    try {
        $__txDir = Join-Path $PSScriptRoot 'Logs'
        if (-not (Test-Path -LiteralPath $__txDir)) { New-Item -ItemType Directory -Path $__txDir -Force | Out-Null }
        $__tx = Join-Path $__txDir ('DeploymentHelper-startup-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        Start-Transcript -LiteralPath $__tx -Force | Out-Null
    } catch { }
}
catch { }

# STA guard: WPF requires STA. Some hosts launch MTA.
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $psExe = (Get-Process -Id $PID).Path
    $fwd = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',$PSCommandPath)
    if (-not [string]::IsNullOrWhiteSpace($SiteCode))    { $fwd += @('-SiteCode',    $SiteCode) }
    if (-not [string]::IsNullOrWhiteSpace($SMSProvider)) { $fwd += @('-SMSProvider', $SMSProvider) }
    Start-Process -FilePath $psExe -ArgumentList $fwd | Out-Null
    try { Stop-Transcript | Out-Null } catch { }
    exit 0
}

# =============================================================================
# Assembly loading
# =============================================================================
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$libDir = Join-Path $PSScriptRoot 'Lib'

Get-ChildItem -LiteralPath $libDir -File -ErrorAction SilentlyContinue |
    Unblock-File -ErrorAction SilentlyContinue

[System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'Microsoft.Xaml.Behaviors.dll')) | Out-Null
[System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'ControlzEx.dll')) | Out-Null
[System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'MahApps.Metro.dll')) | Out-Null

# =============================================================================
# Module import (DeploymentHelperCommon)
# =============================================================================
# Import the shared module. Do NOT SilentlyContinue: if this fails, the
# script would continue but every downstream call to Initialize-Logging /
# Connect-CMSite / Test-ApplicationExists etc. would throw
# CommandNotFoundException at runtime, which crashes the pipeline. Fail
# loudly here instead.
$__modulePath = Join-Path $PSScriptRoot 'Module\DeploymentHelperCommon.psd1'
if (-not (Test-Path -LiteralPath $__modulePath)) {
    throw "Shared module not found at: $__modulePath"
}
Import-Module -Name $__modulePath -Force -DisableNameChecking
if (-not (Get-Command Initialize-Logging -ErrorAction SilentlyContinue)) {
    throw "Shared module imported but Initialize-Logging is not exported. Check DeploymentHelperCommon.psd1."
}

# =============================================================================
# Preferences (DeploymentHelper.prefs.json)
#
# Prefs use $global: scope because they are mutated from inside closures
# (panel Commit scriptblocks that need .GetNewClosure() to capture local
# textbox refs). Per feedback_ps_wpf_handler_rules.md Rule 1 caveat,
# .GetNewClosure() strips access to $script: scope variables, so a
# $script:Prefs read from inside a committed closure resolves to null.
# Global scope survives the stripping and keeps shared mutable state
# reachable across the closure chain.
# =============================================================================
$global:PrefsPath = Join-Path $PSScriptRoot 'DeploymentHelper.prefs.json'

function Get-DhPreferences {
    # Plain hashtable (NOT [ordered]) -- PS 5.1 doesn't expose OrderedDictionary
    # entries as properties, so `$prefs.SiteCode = 'X'` assignments fail.
    $defaults = @{
        SiteCode                = ''
        SMSProvider             = ''
        DeploymentAuditLogPath  = Join-Path $PSScriptRoot 'Logs\deployment-audit.jsonl'
        RingRunStatePath        = Join-Path $PSScriptRoot 'Rings\runs'
    }
    if (Test-Path -LiteralPath $global:PrefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $global:PrefsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            foreach ($k in @($defaults.Keys)) {
                $val = $loaded.$k
                if ($null -ne $val -and -not [string]::IsNullOrWhiteSpace([string]$val)) {
                    $defaults[$k] = $val
                }
            }
        } catch { }
    }
    return $defaults
}

function Save-DhPreferences {
    param([hashtable]$Prefs)
    $null = Save-SuiteSettings -Path $global:PrefsPath -Settings $Prefs
}

$global:Prefs = Get-DhPreferences

# The suite launcher hands its site code and provider to each tool it starts.
# A value saved in this tool wins; the launcher value fills an empty one.
if (-not $global:Prefs.SiteCode    -and $env:SUITE_CM_SITECODE) { $global:Prefs.SiteCode    = [string]$env:SUITE_CM_SITECODE }
if (-not $global:Prefs.SMSProvider -and $env:SUITE_CM_PROVIDER) { $global:Prefs.SMSProvider = [string]$env:SUITE_CM_PROVIDER }

# Script params override prefs only when non-empty. Empty-by-default
# means the shipped app reads from DeploymentHelper.prefs.json (user-
# supplied on first run via Options > Connection) and any CLI override
# overrides that. Never auto-populates a site code or provider the
# user hasn't configured.
if (-not [string]::IsNullOrWhiteSpace($SiteCode))    { $global:Prefs['SiteCode']    = $SiteCode }
if (-not [string]::IsNullOrWhiteSpace($SMSProvider)) { $global:Prefs['SMSProvider'] = $SMSProvider }

# Initialize tool log
$toolLogDir = Join-Path $PSScriptRoot 'Logs'
if (-not (Test-Path -LiteralPath $toolLogDir)) {
    New-Item -ItemType Directory -Path $toolLogDir -Force | Out-Null
}
$toolLogPath = Join-Path $toolLogDir ('DeploymentHelper-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $toolLogPath

# Deployment audit log (JSONL) -- drive off prefs so the Options window can change it
$script:DeploymentAuditLog = $global:Prefs['DeploymentAuditLogPath']
if ([string]::IsNullOrWhiteSpace($script:DeploymentAuditLog)) {
    $script:DeploymentAuditLog = Join-Path $toolLogDir 'deployment-audit.jsonl'
    $global:Prefs['DeploymentAuditLogPath'] = $script:DeploymentAuditLog
}

# =============================================================================
# First-run template seed. Writes four default deployment templates to
# the Templates\ folder if (and only if) the folder is absent or empty.
# Templates the user creates or edits are never overwritten.
# =============================================================================
$__templatesDir = Join-Path $PSScriptRoot 'Templates'
if (-not (Test-Path -LiteralPath $__templatesDir)) {
    New-Item -ItemType Directory -Path $__templatesDir -Force | Out-Null
}
if (-not (Get-ChildItem -LiteralPath $__templatesDir -Filter '*.json' -ErrorAction SilentlyContinue)) {
    $__seed = @(
        [ordered]@{
            Name                       = 'Workstation Pilot'
            DeployPurpose              = 'Available'
            UserNotification           = 'DisplaySoftwareCenterOnly'
            TimeBasedOn                = 'LocalTime'
            OverrideServiceWindow      = $false
            RebootOutsideServiceWindow = $false
            AllowMeteredConnection     = $false
            AllowBoundaryFallback      = $true
            AllowMicrosoftUpdate       = $false
            RequirePostRebootFullScan  = $true
            DefaultDeadlineOffsetHours = 0
        },
        [ordered]@{
            Name                       = 'Workstation Production'
            DeployPurpose              = 'Required'
            UserNotification           = 'DisplayAll'
            TimeBasedOn                = 'LocalTime'
            OverrideServiceWindow      = $false
            RebootOutsideServiceWindow = $false
            AllowMeteredConnection     = $true
            AllowBoundaryFallback      = $true
            AllowMicrosoftUpdate       = $false
            RequirePostRebootFullScan  = $true
            DefaultDeadlineOffsetHours = 72
        },
        [ordered]@{
            Name                       = 'Server Pilot'
            DeployPurpose              = 'Available'
            UserNotification           = 'DisplaySoftwareCenterOnly'
            TimeBasedOn                = 'LocalTime'
            OverrideServiceWindow      = $false
            RebootOutsideServiceWindow = $false
            AllowMeteredConnection     = $false
            AllowBoundaryFallback      = $true
            AllowMicrosoftUpdate       = $false
            RequirePostRebootFullScan  = $true
            DefaultDeadlineOffsetHours = 0
        },
        [ordered]@{
            Name                       = 'Server Production'
            DeployPurpose              = 'Required'
            UserNotification           = 'HideAll'
            TimeBasedOn                = 'LocalTime'
            OverrideServiceWindow      = $false
            RebootOutsideServiceWindow = $false
            AllowMeteredConnection     = $true
            AllowBoundaryFallback      = $true
            AllowMicrosoftUpdate       = $false
            RequirePostRebootFullScan  = $true
            DefaultDeadlineOffsetHours = 168
        }
    )
    foreach ($t in $__seed) {
        $__fname = ($t.Name -replace '\s+','') + '.json'
        $__fpath = Join-Path $__templatesDir $__fname
        try {
            $t | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $__fpath -Encoding UTF8
        } catch { }
    }
}

# =============================================================================
# Ring plans: Rings\*.json beside Templates\. First run writes two seed plans
# that name no collection. Run state for held rings lives in the folder set
# in Options > Logging (default Rings\runs; a UNC path lets a team share it).
# =============================================================================
$script:RingPlanDir = Join-Path $PSScriptRoot 'Rings'
try { [void](Initialize-RingPlanFolder -Path $script:RingPlanDir) }
catch { Write-Log ('Ring plan seed failed: {0}' -f $_.Exception.Message) -Level WARN }
$script:RingRunStatePath = [string]$global:Prefs['RingRunStatePath']

# =============================================================================
# Deployment type state
# =============================================================================
$script:CurrentType = 'Apps'
$script:ConnectedToCM = $false

$script:ValidatedApp        = $null
$script:ValidatedCollection = $null
$script:AllChecksPassed     = $false

$script:ValidatedPackage      = $null
$script:ValidatedProgram      = $null
$script:ValidatedTaskSequence = $null
$script:ValidatedSUG          = $null

# DP group cache + per-type selection state (global: for closure safety)
$global:DPGroupsCache    = $null   # string[] of DP group names; $null = not loaded
$global:SelectedDPGroups = @()     # names currently selected in the picker

$script:TypeMeta = @{
    'Apps' = @{
        Header      = 'Apps Deployment'
        Subheader   = 'Select an application and a device collection, then configure the deployment.'
        TargetLabel = 'Application:'
        Watermark   = 'e.g. 7-Zip 26.00'
        Check1Text  = 'Application exists in ConfigMgr'
    }
    'Packages' = @{
        Header      = 'Packages Deployment'
        Subheader   = 'Select a classic package, program, and device collection.'
        TargetLabel = 'Package:'
        Watermark   = 'Classic package name'
        Check1Text  = 'Package and program exist in ConfigMgr'
    }
    'TaskSequences' = @{
        Header      = 'Task Sequences Deployment'
        Subheader   = 'Select a task sequence and a device collection.'
        TargetLabel = 'Task Sequence:'
        Watermark   = 'Task sequence name'
        Check1Text  = 'Task sequence exists in ConfigMgr'
    }
    'SUG' = @{
        Header      = 'Software Update Groups Deployment'
        Subheader   = 'Select a software update group and a device collection.'
        TargetLabel = 'Update Group:'
        Watermark   = 'Software update group name'
        Check1Text  = 'Software update group exists in ConfigMgr'
    }
}

# =============================================================================
# Log drawer and status helpers
# =============================================================================
function Add-LogLine {
    param([Parameter(Mandatory)][string]$Message)

    $ts = (Get-Date).ToString('HH:mm:ss')
    $line = '{0}  {1}' -f $ts, $Message

    if ([string]::IsNullOrWhiteSpace($txtLog.Text)) {
        $txtLog.Text = $line
    }
    else {
        $txtLog.AppendText([Environment]::NewLine + $line)
    }
    $txtLog.ScrollToEnd()
}

function Set-StatusText {
    param([Parameter(Mandatory)][string]$Text)
    $txtStatus.Text = $Text
}

function Show-AuditWriteFailure {
    param(
        [Parameter(Mandatory)][hashtable]$AuditResult,
        [bool]$DeploymentCreated,
        [string]$DeploymentID
    )

    $message = if ($DeploymentCreated) {
        "Configuration Manager created deployment $DeploymentID, but the audit record could not be written to $script:DeploymentAuditLog. Record the deployment ID and restore audit logging before continuing. Error: $($AuditResult.Error)"
    } else {
        "The deployment attempt could not be written to the audit log at $script:DeploymentAuditLog. Error: $($AuditResult.Error)"
    }
    Add-LogLine -Message $message
    Set-StatusText -Text $(if ($DeploymentCreated) { 'Deployment created; audit write failed.' } else { 'Audit write failed.' })
    [void](Show-ThemedMessage -Owner $window -Title 'Audit log write failed' -Message $message -Buttons OK -Icon Warn)
}

function Test-DeploymentAuditReady {
    $check = Test-DeploymentLogWritable -LogPath $script:DeploymentAuditLog
    if ($check.Success) { return $true }

    $message = ("Deployment was blocked because the audit log is not writable at {0}: {1}" -f $script:DeploymentAuditLog, $check.Error)
    Add-LogLine -Message $message
    Set-StatusText -Text 'Deployment blocked: audit log unavailable.'
    [void](Show-ThemedMessage -Owner $window -Title 'Deployment blocked' -Message $message -Buttons OK -Icon Error)
    return $false
}

function Test-NoCurrentDuplicateDeployment {
    param(
        [Parameter(Mandatory)][ValidateSet('Application', 'Package', 'TaskSequence', 'SUG')][string]$Type,
        [Parameter(Mandatory)]$TargetObject,
        [Parameter(Mandatory)]$Collection,
        [string]$ProgramName
    )

    $duplicate = switch ($Type) {
        'Application'  { Test-DuplicateDeployment -ApplicationName $TargetObject.LocalizedDisplayName -CollectionName $Collection.Name }
        'Package'      { Test-DuplicatePackageDeployment -PackageID $TargetObject.PackageID -ProgramName $ProgramName -CollectionName $Collection.Name }
        'TaskSequence' { Test-DuplicateTaskSequenceDeployment -TaskSequencePackageId $TargetObject.PackageID -CollectionName $Collection.Name }
        'SUG'          { Test-DuplicateSUGDeployment -SUGName $TargetObject.LocalizedDisplayName -CollectionName $Collection.Name }
    }
    $checkFailure = Get-DuplicateCheckFailure -Result $duplicate
    if ($checkFailure) {
        $message = ("Could not recheck for a duplicate deployment: {0}. No deployment was created." -f $checkFailure.Error)
    }
    elseif ($null -ne $duplicate) {
        $message = ("A matching deployment now exists on '{0}'. No duplicate was created." -f $Collection.Name)
    }
    else {
        return $true
    }

    Add-LogLine -Message $message
    Set-StatusText -Text 'Deployment blocked by the final duplicate check.'
    [void](Show-ThemedMessage -Owner $window -Title 'Deployment blocked' -Message $message -Buttons OK -Icon Warn)
    return $false
}

# =============================================================================
# Window state persistence (geometry logic: SuiteCommon)
# =============================================================================
function Get-WindowStatePath {
    Join-Path $PSScriptRoot 'DeploymentHelper.windowstate.json'
}

# =============================================================================
# Parse XAML and create window
# =============================================================================
$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
[xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Installed version: the script header is the single source of truth for the
# About panel.
$script:AppVersion = '0.0.0'
foreach ($headerLine in (Get-Content -LiteralPath $PSCommandPath -TotalCount 80)) {
    if ($headerLine -match '^\s*Version\s*:\s*([0-9][0-9\.]*[0-9])\s*$') { $script:AppVersion = $Matches[1]; break }
}

# =============================================================================
# Title-bar drag fallback. PS51-WPF-033. SuiteCommon owns the hook.
# =============================================================================
Install-TitleBarDragFallback -Window $window

# =============================================================================
# Find named controls
# =============================================================================
$toggleTheme        = $window.FindName('toggleTheme')
$txtThemeLabel      = $window.FindName('txtThemeLabel')
$btnApps            = $window.FindName('btnApps')
$btnPackages        = $window.FindName('btnPackages')
$btnTaskSequences   = $window.FindName('btnTaskSequences')
$btnSUG             = $window.FindName('btnSUG')
$btnOptions         = $window.FindName('btnOptions')
$txtModuleHeader    = $window.FindName('txtModuleHeader')
$txtModuleSubheader = $window.FindName('txtModuleSubheader')
$lblTargetName      = $window.FindName('lblTargetName')
$txtTargetName      = $window.FindName('txtTargetName')
$btnBrowseTarget    = $window.FindName('btnBrowseTarget')
$txtCollection      = $window.FindName('txtCollection')
$btnBrowseCollection = $window.FindName('btnBrowseCollection')
$radAvailable       = $window.FindName('radAvailable')
$radRequired        = $window.FindName('radRequired')
$dtpAvailable       = $window.FindName('dtpAvailable')
$lblDeadline        = $window.FindName('lblDeadline')
$dtpDeadline        = $window.FindName('dtpDeadline')
$cboNotification    = $window.FindName('cboNotification')
$radLocalTime       = $window.FindName('radLocalTime')
$radUtc             = $window.FindName('radUtc')
$pnlRequiredOptions = $window.FindName('pnlRequiredOptions')
$chkOverrideSW      = $window.FindName('chkOverrideSW')
$chkRebootOutSW     = $window.FindName('chkRebootOutSW')
$chkMetered         = $window.FindName('chkMetered')
$lblProgram         = $window.FindName('lblProgram')
$cboProgram         = $window.FindName('cboProgram')
$lblNotification    = $window.FindName('lblNotification')
$pnlPackageOptions  = $window.FindName('pnlPackageOptions')
$chkPkgOverrideSW   = $window.FindName('chkPkgOverrideSW')
$chkPkgRebootOutSW  = $window.FindName('chkPkgRebootOutSW')
$chkPkgMetered      = $window.FindName('chkPkgMetered')
$cboPkgFastNetwork  = $window.FindName('cboPkgFastNetwork')
$cboPkgSlowNetwork  = $window.FindName('cboPkgSlowNetwork')
$cboPkgRerun        = $window.FindName('cboPkgRerun')
$pnlTaskSequenceOptions = $window.FindName('pnlTaskSequenceOptions')
$chkTsOverrideSW    = $window.FindName('chkTsOverrideSW')
$chkTsRebootOutSW   = $window.FindName('chkTsRebootOutSW')
$chkTsMetered       = $window.FindName('chkTsMetered')
$chkShowProgress    = $window.FindName('chkShowProgress')
$pnlSUGOptions      = $window.FindName('pnlSUGOptions')
$chkSugOverrideSW   = $window.FindName('chkSugOverrideSW')
$chkSugAllowRestart = $window.FindName('chkSugAllowRestart')
$chkSugMetered      = $window.FindName('chkSugMetered')
$chkMSFallback      = $window.FindName('chkMSFallback')
$chkBoundaryFallback = $window.FindName('chkBoundaryFallback')
$chkFullScan        = $window.FindName('chkFullScan')
$lblDistribute      = $window.FindName('lblDistribute')
$pnlDistribute      = $window.FindName('pnlDistribute')
$btnDPPicker        = $window.FindName('btnDPPicker')
$btnDistributeContent = $window.FindName('btnDistributeContent')
$lblCheck1          = $window.FindName('lblCheck1')
$glyphCheck1        = $window.FindName('glyphCheck1')
$glyphCheck2        = $window.FindName('glyphCheck2')
$glyphCheck3        = $window.FindName('glyphCheck3')
$glyphCheck4        = $window.FindName('glyphCheck4')
$glyphCheck5        = $window.FindName('glyphCheck5')
$btnValidate        = $window.FindName('btnValidate')
$btnDeploy          = $window.FindName('btnDeploy')
$txtLog             = $window.FindName('txtLog')
$lblLogOutput       = $window.FindName('lblLogOutput')
$txtStatus          = $window.FindName('txtStatus')
$cboApplyTemplate   = $window.FindName('cboApplyTemplate')
$pnlApplyTemplate   = $window.FindName('pnlApplyTemplate')
$formPane           = $window.FindName('formPane')
$btnRings           = $window.FindName('btnRings')
$ringPane           = $window.FindName('ringPane')
$cboRingType        = $window.FindName('cboRingType')
$txtRingObject      = $window.FindName('txtRingObject')
$btnRingBrowseObject = $window.FindName('btnRingBrowseObject')
$lblRingProgram     = $window.FindName('lblRingProgram')
$cboRingProgram     = $window.FindName('cboRingProgram')
$cboRingPlan        = $window.FindName('cboRingPlan')
$btnRingReloadPlans = $window.FindName('btnRingReloadPlans')
$btnRingOpenFolder  = $window.FindName('btnRingOpenFolder')
$dtpRingStart       = $window.FindName('dtpRingStart')
$btnRingExpand      = $window.FindName('btnRingExpand')
$chkRingHold        = $window.FindName('chkRingHold')
$txtRingPlanInfo    = $window.FindName('txtRingPlanInfo')
$dgRingPreview      = $window.FindName('dgRingPreview')
$btnRingPickCollection = $window.FindName('btnRingPickCollection')
$btnRingValidate    = $window.FindName('btnRingValidate')
$btnRingCreate      = $window.FindName('btnRingCreate')
$lblRingRuns        = $window.FindName('lblRingRuns')
$btnRingRefreshRuns = $window.FindName('btnRingRefreshRuns')
$chkRingShowClosed  = $window.FindName('chkRingShowClosed')
$dgRingRuns         = $window.FindName('dgRingRuns')

# Seed datetime pickers so SelectedDate is non-null at first render
$dtpAvailable.SelectedDateTime = Get-Date
$dtpDeadline.SelectedDateTime  = (Get-Date).AddHours(24)
$__now = Get-Date
$dtpRingStart.SelectedDateTime = $__now.Date.AddHours($__now.Hour).AddMinutes(($__now.Minute - ($__now.Minute % 5)) + 5)

# =============================================================================
# Theme setup and toggle
# =============================================================================
[void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Dark.Steel')

$script:WorkflowButtons = @($btnApps, $btnPackages, $btnTaskSequences, $btnSUG, $btnRings)
$script:OptionsButtons  = @($btnOptions)

# Initialize-SuiteTheme must run before the first Set-ButtonTheme call
# (SourceInitialized handler below, and the toggle's Add_Toggled) --
# Set-ButtonTheme throws if the module's theme context is still empty.
Initialize-SuiteTheme -Window $window `
    -IsDarkGetter { $toggleTheme.IsOn -eq $true } `
    -WorkflowButtons $script:WorkflowButtons `
    -OptionsButtons $script:OptionsButtons `
    -LogLabel $lblLogOutput

$script:TitleBarBlue         = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:TitleBarBlueInactive = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#4BA3E0')

# LOG OUTPUT label Foreground per theme. Single hex fails AA on one theme.
# See reference_srl_wpf_brand.md.
$script:LogLabelDark  = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#B0B0B0')
$script:LogLabelLight = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#595959')

$toggleTheme.Add_Toggled({
    if ($toggleTheme.IsOn) {
        [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Dark.Steel')
        $txtThemeLabel.Text = 'Dark Theme'
        Set-ButtonTheme -IsDark $true
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::WindowTitleBrushProperty)
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::NonActiveWindowTitleBrushProperty)
    }
    else {
        [ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Light.Blue')
        $txtThemeLabel.Text = 'Light Theme'
        Set-ButtonTheme -IsDark $false
        $window.WindowTitleBrush         = $script:TitleBarBlue
        $window.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }
})

# =============================================================================
# Validation pane helpers
# =============================================================================
$script:GlyphPending = [char]0x22EF  # horizontal ellipsis
$script:GlyphPass    = [char]0x2713  # check mark
$script:GlyphFail    = [char]0x2717  # ballot x

function Set-CheckGlyph {
    # Per feedback_no_red_green_in_brand.md, state is carried by the glyph
    # shape only. Pending = horizontal ellipsis, Pass = checkmark, Fail =
    # ballot-x. Foreground is NEVER set here -- it stays at ThemeForeground
    # via inheritance so both themes render at AAA contrast.
    param(
        [Parameter(Mandatory)][ValidateRange(1,5)][int]$Index,
        [Parameter(Mandatory)][ValidateSet('Pending','Pass','Fail')][string]$State
    )
    $glyph = switch ($State) {
        'Pending' { $script:GlyphPending }
        'Pass'    { $script:GlyphPass }
        'Fail'    { $script:GlyphFail }
    }
    $glyphControl = switch ($Index) {
        1 { $glyphCheck1 }
        2 { $glyphCheck2 }
        3 { $glyphCheck3 }
        4 { $glyphCheck4 }
        5 { $glyphCheck5 }
    }
    $glyphControl.Text = [string]$glyph
}

function Reset-ValidationUi {
    for ($i = 1; $i -le 5; $i++) { Set-CheckGlyph -Index $i -State 'Pending' }
    $script:ValidatedApp          = $null
    $script:ValidatedCollection   = $null
    $script:ValidatedPackage      = $null
    $script:ValidatedProgram      = $null
    $script:ValidatedTaskSequence = $null
    $script:ValidatedSUG          = $null
    $script:AllChecksPassed       = $false
    $btnDeploy.Visibility = [System.Windows.Visibility]::Collapsed
}

# =============================================================================
# Notification/Availability dropdown item sets
# =============================================================================
$script:AppsNotificationItems = @(
    @{ Content = 'Display All';                      Tag = 'DisplayAll' },
    @{ Content = 'Display in Software Center Only';  Tag = 'DisplaySoftwareCenterOnly' },
    @{ Content = 'Hide All';                         Tag = 'HideAll' }
)
$script:TaskSequenceAvailabilityItems = @(
    @{ Content = 'Clients';                  Tag = 'Clients' },
    @{ Content = 'Clients, media, and PXE';  Tag = 'ClientsMediaAndPxe' },
    @{ Content = 'Media and PXE';            Tag = 'MediaAndPxe' },
    @{ Content = 'Media and PXE (hidden)';   Tag = 'MediaAndPxeHidden' }
)

function Set-NotificationItems {
    param([array]$Items, [string]$LabelText)
    $cboNotification.Items.Clear()
    foreach ($i in $Items) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $i.Content
        $item.Tag     = $i.Tag
        [void]$cboNotification.Items.Add($item)
    }
    if ($cboNotification.Items.Count -gt 0) { $cboNotification.SelectedIndex = 0 }
    $lblNotification.Text = $LabelText
}

# =============================================================================
# CM site connect helper
# =============================================================================
function Connect-IfNeeded {
    if ($script:ConnectedToCM -and (Test-CMConnection)) { return $true }

    $sc   = [string]$global:Prefs['SiteCode']
    $smsp = [string]$global:Prefs['SMSProvider']
    if ([string]::IsNullOrWhiteSpace($sc) -or [string]::IsNullOrWhiteSpace($smsp)) {
        Add-LogLine -Message 'Site code and SMS provider not set. Open Options > Connection.'
        Set-StatusText -Text 'Connection not configured.'
        return $false
    }

    Add-LogLine -Message ('Connecting to CM site {0} on {1}...' -f $sc, $smsp)
    Set-StatusText -Text 'Connecting...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        $ok = Connect-CMSite -SiteCode $sc -SMSProvider $smsp
    }
    finally {
        $window.Cursor = $null
    }
    if ($ok) {
        $script:ConnectedToCM = $true
        Add-LogLine -Message 'Connected.'
        Set-StatusText -Text ('Connected to {0}.' -f $sc)
        return $true
    }
    else {
        $script:ConnectedToCM = $false
        Add-LogLine -Message 'Failed to connect. Check site code, provider, and ConfigurationManager module.'
        Set-StatusText -Text 'Connection failed.'
        return $false
    }
}

# =============================================================================
# Themed message dialog (replaces System.Windows.MessageBox::Show calls so
# every confirmation/error stays within the MahApps theme): SuiteCommon
# =============================================================================

# =============================================================================
# Deployment schedule sanity checks
# Returns @{ Ok = $bool; Level = 'Error'|'Warn'|'Info'; Reason = <string> }.
# - Deadline <= Available for a Required deploy is nonsensical. Block.
# - Deadline in the past for a Required deploy will fire immediately. Block.
# - Available more than 5 minutes in the past is almost always a backdate
#   mistake. Warn + confirm, not block (could be intentional).
# - Available more than 5 years in the future is almost certainly typo. Warn.
# =============================================================================
function Test-DeploymentNotification {
    # Guards the "Available + HideAll" combo. Both halves are individually
    # valid, together they produce a deployment that neither surfaces to
    # the user in Software Center NOR auto-installs -- a silent no-op.
    # ConfigMgr accepts it; we don't.
    param(
        [Parameter(Mandatory)][ValidateSet('Available','Required')][string]$Purpose,
        [Parameter(Mandatory)][ValidateSet('DisplayAll','DisplaySoftwareCenterOnly','HideAll')][string]$Notification
    )
    if ($Purpose -eq 'Available' -and $Notification -eq 'HideAll') {
        return @{
            Ok     = $false
            Level  = 'Error'
            Reason = 'Notification = Hide All is incompatible with Purpose = Available: the deployment would be invisible to users AND never auto-install. Change one.'
        }
    }
    return @{ Ok = $true; Level = 'Info'; Reason = '' }
}

function Test-DeploymentSchedule {
    param(
        [Parameter(Mandatory)][datetime]$Available,
        # Nullable: Available deployments legitimately pass $null here
        # (dtpDeadline.SelectedDateTime is $null when the user hasn't
        # opened the deadline picker). A bare [datetime] parameter binds
        # $null to System.DateTime and throws
        # "Cannot convert null to type 'System.DateTime'" before the
        # function body runs -- the null-check on the next line never
        # gets a chance to fire.
        [Nullable[datetime]]$Deadline,
        [Parameter(Mandatory)][ValidateSet('Available','Required')][string]$Purpose,
        # With time basis UTC the pickers hold UTC wall-clock values; the
        # caller passes UtcNow so a passed UTC deadline is still caught.
        [datetime]$Now = (Get-Date)
    )

    $now = $Now

    if ($Purpose -eq 'Required') {
        if ($null -eq $Deadline) {
            return @{ Ok = $false; Level = 'Error'; Reason = 'Deadline is required when Purpose is Required.' }
        }
        if ($Deadline -le $Available) {
            return @{ Ok = $false; Level = 'Error';
                Reason = ("Deadline ({0}) must be after Available ({1}). A Required deployment with a deadline on or before its available time would fire immediately." -f
                    $Deadline.ToString('yyyy-MM-dd HH:mm'), $Available.ToString('yyyy-MM-dd HH:mm')) }
        }
        if ($Deadline -le $now) {
            return @{ Ok = $false; Level = 'Error';
                Reason = ("Deadline ({0}) is in the past. A Required deployment with a past deadline would fire immediately on every targeted client." -f
                    $Deadline.ToString('yyyy-MM-dd HH:mm')) }
        }
    }

    if ($Available -lt $now.AddMinutes(-5)) {
        return @{ Ok = $true; Level = 'Warn';
            Reason = ("Available time ({0}) is in the past. Clients will see the deployment immediately. Continue anyway?" -f
                $Available.ToString('yyyy-MM-dd HH:mm')) }
    }

    if ($Available -gt $now.AddYears(5)) {
        return @{ Ok = $true; Level = 'Warn';
            Reason = ("Available time ({0}) is more than five years away. Is this correct?" -f
                $Available.ToString('yyyy-MM-dd HH:mm')) }
    }

    return @{ Ok = $true; Level = 'Info'; Reason = '' }
}

# =============================================================================
# Background runspace for browse loads
# =============================================================================
$script:BgRunspace       = $null
$script:BgPowerShell     = $null
$script:BgInvokeHandle   = $null
$script:BgTimer          = $null
$script:BgGraveyard      = @()
$script:BrowseLoadState  = $null
$script:BrowseLoadDialog = $null
$script:BrowseLoadResult = $null
$script:BrowseCache      = @{}

function Initialize-BgRunspace {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Lazy-init of the background runspace; idempotent.')]
    param()
    if ($script:BgRunspace -and $script:BgRunspace.RunspaceStateInfo.State -eq 'Opened') { return }
    $script:BgRunspace = New-SuiteBgRunspace -ModulePath $__modulePath -LogPath $toolLogPath
}

function Stop-BgWork {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Tears down ephemeral runspace plumbing only.')]
    param()
    $script:BgGraveyard = @(Stop-SuiteBgWork -PowerShell $script:BgPowerShell -Timer $script:BgTimer -Graveyard $script:BgGraveyard)
    $script:BgTimer        = $null
    $script:BgPowerShell   = $null
    $script:BgInvokeHandle = $null
}

# =============================================================================
# Browse loading dialog: runs one bulk read off the dispatcher and returns
# @{ Items; Folders } or $null (cancel / error).
# =============================================================================
function Show-BrowseLoadingDialog {
    param(
        [Parameter(Mandatory)]$Owner,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][ValidateSet('Apps', 'Packages', 'TaskSequences', 'SUG', 'Collections')][string]$Type
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Loading"
    Width="420" Height="170"
    ResizeMode="NoResize"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    ShowCloseButton="False"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <StackPanel Grid.Row="0" Orientation="Horizontal" VerticalAlignment="Center">
            <Controls:ProgressRing IsActive="True" Width="36" Height="36" Margin="0,0,16,0"
                                   Foreground="{DynamicResource MahApps.Brushes.Accent}"/>
            <TextBlock x:Name="txtStep" VerticalAlignment="Center" FontSize="13" Text="Connecting..."/>
        </StackPanel>
        <Button x:Name="btnCancel" Grid.Row="1" Content="Cancel" MinWidth="90" Height="32" HorizontalAlignment="Right" IsCancel="True"
                Style="{DynamicResource MahApps.Styles.Button.Square}"
                Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$xml = $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $dlg    = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg
    $dlg.Owner = $Owner
    $dlg.Title = $Title
    Set-DialogTheme -Dialog $dlg

    $btnCancel = $dlg.FindName('btnCancel')

    # New-SuiteBgRunspace throws when the module bootstrap fails; an
    # uncaught throw here takes down the dispatcher.
    try {
        Initialize-BgRunspace
    }
    catch {
        Add-LogLine -Message ('Browse load failed: {0}' -f $_.Exception.Message)
        [void](Show-ThemedMessage -Owner $Owner -Title 'Browse' -Message ('Background load could not start: {0}' -f $_.Exception.Message) -Buttons OK -Icon Error)
        return $null
    }
    Stop-BgWork

    $script:BrowseLoadDialog = $dlg
    $script:BrowseLoadResult = $null
    $script:BrowseLoadState  = [hashtable]::Synchronized(@{
        Step     = 'Connecting...'
        Done     = $false
        Result   = $null
        ErrorMsg = $null
    })

    $siteCode    = [string]$global:Prefs['SiteCode']
    $smsProvider = [string]$global:Prefs['SMSProvider']

    $script:BgPowerShell = [powershell]::Create()
    $script:BgPowerShell.Runspace = $script:BgRunspace
    [void]$script:BgPowerShell.AddScript({
        param($SiteCode, $SMSProvider, $Type, $State)
        try {
            # The CM site drive and current location are per runspace; the
            # UI-thread connection does not reach this runspace.
            if (-not (Test-CMConnection)) {
                $State.Step = "Connecting to $SiteCode..."
                if (-not (Connect-CMSite -SiteCode $SiteCode -SMSProvider $SMSProvider)) {
                    $State.ErrorMsg = "Failed to connect to site $SiteCode (provider $SMSProvider)."
                    return
                }
            }

            $label = switch ($Type) {
                'Apps'          { 'applications' }
                'Packages'      { 'packages' }
                'TaskSequences' { 'task sequences' }
                'SUG'           { 'software update groups' }
                'Collections'   { 'device collections' }
            }
            $State.Step = "Loading $label..."
            $items   = @(Get-CMBrowseList -Type $Type)
            $folders = @()

            if ($Type -eq 'Collections') {
                $State.Step = 'Loading collection folders...'
                $folderMap = @{}
                try {
                    $info      = Get-CMCollectionFolderInfo -SMSProvider $SMSProvider -SiteCode $SiteCode
                    $folders   = @($info.Folders)
                    $folderMap = $info.FolderMap
                } catch {
                    # A blocked CIM read degrades the picker to a flat root list.
                    Write-Log ("Collection folder read failed; showing a flat list: {0}" -f $_.Exception.Message) -Level WARN
                    $folders   = @()
                    $folderMap = @{}
                }
                $items = @(Add-CollectionFolderId -Collections $items -FolderMap $folderMap)
            }

            $State.Result = [PSCustomObject]@{ Items = $items; Folders = $folders }
        }
        catch {
            $State.ErrorMsg = $_.Exception.Message
        }
        finally {
            $State.Done = $true
        }
    }).AddArgument($siteCode).AddArgument($smsProvider).AddArgument($Type).AddArgument($script:BrowseLoadState)

    $script:BgInvokeHandle = $script:BgPowerShell.BeginInvoke()

    # $script: state only in the tick handler: the timer outlives the
    # dialog's Loaded scope and the handler must not depend on lexical refs.
    $script:BgTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:BgTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:BgTimer.Add_Tick({
        $state = $script:BrowseLoadState
        if (-not $state) { return }
        $step = [string]$state.Step
        $lbl  = $script:BrowseLoadDialog.FindName('txtStep')
        if ($lbl -and $lbl.Text -ne $step) { $lbl.Text = $step }
        if (-not $state.Done) { return }

        $script:BgTimer.Stop()
        try { [void]$script:BgPowerShell.EndInvoke($script:BgInvokeHandle) } catch { $null = $_ }
        try { $script:BgPowerShell.Dispose() } catch { $null = $_ }
        $script:BgPowerShell   = $null
        $script:BgInvokeHandle = $null

        if ($state.ErrorMsg) {
            $script:BrowseLoadResult = $null
            Add-LogLine -Message ('Browse load failed: {0}' -f $state.ErrorMsg)
            [void](Show-ThemedMessage -Owner $script:BrowseLoadDialog -Title 'Browse' -Message ('Load error: {0}' -f $state.ErrorMsg) -Buttons OK -Icon Error)
        }
        else {
            $script:BrowseLoadResult = $state.Result
        }
        $script:BrowseLoadDialog.Close()
    })

    $btnCancel.Add_Click({
        $script:BrowseLoadResult = $null
        Stop-BgWork
        $dlg.Close()
    })

    # Alt+F4 bypasses the Cancel button; the pipeline must not outlive the dialog.
    $dlg.Add_Closing({ Stop-BgWork })

    $dlg.Add_Loaded({ $script:BgTimer.Start() })
    [void]$dlg.ShowDialog()

    $script:BrowseLoadDialog = $null
    $result = $script:BrowseLoadResult
    $script:BrowseLoadResult = $null
    return $result
}

function Get-BrowseList {
    <#
    .SYNOPSIS
        Returns the cached rows for one object type, loading them once per
        session. -Force reloads. $null when the load was cancelled or failed.
    #>
    param(
        [Parameter(Mandatory)]$Owner,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Type,
        [switch]$Force
    )
    if (-not $Force -and $script:BrowseCache.ContainsKey($Type)) { return $script:BrowseCache[$Type] }
    $loaded = Show-BrowseLoadingDialog -Owner $Owner -Title $Title -Type $Type
    if ($null -eq $loaded) { return $null }
    $script:BrowseCache[$Type] = $loaded
    return $loaded
}

# =============================================================================
# Themed browse dialog (Application / Package / Task Sequence / SUG)
# =============================================================================
function Show-BrowseDialog {
    param(
        [Parameter(Mandatory)]$Owner,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Watermark,
        [Parameter(Mandatory)][ValidateSet('Apps', 'Packages', 'TaskSequences', 'SUG')][string]$Type,
        [Parameter(Mandatory)][string]$NameProperty
    )

    $loaded = Get-BrowseList -Owner $Owner -Title $Title -Type $Type
    if ($null -eq $loaded) { return $null }

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Browse"
    Width="760" Height="520"
    MinWidth="540" MinHeight="360"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <TextBox x:Name="txtFilter" Grid.Row="0" Grid.Column="0" FontSize="12" Height="28" VerticalContentAlignment="Center" Margin="0,0,8,8"/>
        <Button  x:Name="btnRefresh" Grid.Row="0" Grid.Column="1" Content="Refresh" Width="90" Height="28" Margin="0,0,0,8"
                 Style="{DynamicResource MahApps.Styles.Button.Square}"
                 Controls:ControlsHelper.ContentCharacterCasing="Normal"/>

        <DataGrid x:Name="dgResults" Grid.Row="1" Grid.ColumnSpan="2"
                  AutoGenerateColumns="False"
                  CanUserAddRows="False"
                  CanUserDeleteRows="False"
                  IsReadOnly="True"
                  SelectionMode="Single"
                  SelectionUnit="FullRow"
                  GridLinesVisibility="Horizontal"
                  HeadersVisibility="Column"
                  RowHeaderWidth="0"
                  BorderThickness="0"
                  ColumnHeaderHeight="30"
                  EnableRowVirtualization="True"
                  VirtualizingPanel.VirtualizationMode="Recycling"/>

        <TextBlock x:Name="txtStatus" Grid.Row="2" Grid.ColumnSpan="2" FontSize="11" Margin="0,8,0,0"
                   Foreground="{DynamicResource MahApps.Brushes.Gray1}" Text=""/>

        <StackPanel Grid.Row="3" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
            <Button x:Name="btnOK"     Content="OK"     MinWidth="90" Height="32" Margin="0,0,8,0" IsDefault="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnCancel" Content="Cancel" MinWidth="90" Height="32" IsCancel="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$xml = $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $dlg    = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg

    $theme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
    if ($theme) { [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $theme) }
    $dlg.Owner = $Owner
    try {
        $dlg.WindowTitleBrush          = $Owner.WindowTitleBrush
        $dlg.NonActiveWindowTitleBrush = $Owner.WindowTitleBrush
        $dlg.GlowBrush                 = $Owner.GlowBrush
        $dlg.NonActiveGlowBrush        = $Owner.GlowBrush
    } catch { }

    $dlg.Title  = $Title
    $txtFilter  = $dlg.FindName('txtFilter')
    $btnRefresh = $dlg.FindName('btnRefresh')
    $dgResults  = $dlg.FindName('dgResults')
    $txtStatus  = $dlg.FindName('txtStatus')
    $btnOK      = $dlg.FindName('btnOK')
    $btnCancel  = $dlg.FindName('btnCancel')
    [MahApps.Metro.Controls.TextBoxHelper]::SetWatermark($txtFilter, $Watermark)

    # Shared mutable rows live in a hashtable so the handlers below can
    # replace the list after a Refresh without a $script: variable.
    $state = @{ All = @($loaded.Items) }

    # No .GetNewClosure() on any handler in this modal dialog: ShowDialog()
    # blocks this function while the handlers fire, so lexical parent scope
    # reaches $dlg / $dgResults / $state / $NameProperty. GetNewClosure would
    # strip script-function lookup (Show-ThemedMessage, Select-BrowseMatch)
    # and $script: writes.
    $applyFilter = {
        $rows = Select-BrowseMatch -Items $state.All -Needle ([string]$txtFilter.Text)
        $dgResults.ItemsSource = $rows
        $total = @($state.All).Count
        if (@($rows).Count -eq $total) { $txtStatus.Text = ('{0} item(s). Type to filter, or pick a row.' -f $total) }
        else                            { $txtStatus.Text = ('{0} of {1} item(s) match.' -f @($rows).Count, $total) }
    }

    $buildColumns = {
        $dgResults.Columns.Clear()
        if (@($state.All).Count -eq 0) { return }
        $props = @($state.All)[0].PSObject.Properties.Name
        foreach ($p in $props) {
            $col = New-Object System.Windows.Controls.DataGridTextColumn
            $col.Header  = $p
            $col.Binding = New-Object System.Windows.Data.Binding($p)
            if ($p -eq $props[0]) { $col.Width = [System.Windows.Controls.DataGridLength]::new(1, [System.Windows.Controls.DataGridLengthUnitType]::Star) }
            else                  { $col.Width = [System.Windows.Controls.DataGridLength]::SizeToCells }
            $dgResults.Columns.Add($col) | Out-Null
        }
    }

    & $buildColumns
    & $applyFilter

    $txtFilter.Add_TextChanged({ & $applyFilter })

    $btnRefresh.Add_Click({
        $fresh = Get-BrowseList -Owner $dlg -Title $Title -Type $Type -Force
        if ($null -eq $fresh) { return }
        $state.All = @($fresh.Items)
        & $buildColumns
        & $applyFilter
    })

    $dgResults.Add_MouseDoubleClick({
        if ($dgResults.SelectedItem) { $btnOK.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }
    })

    $script:DialogResult = $null
    $btnOK.Add_Click({
        if ($dgResults.SelectedItem) {
            $item = $dgResults.SelectedItem
            $script:DialogResult = [string]$item.$NameProperty
            $dlg.Close()
        }
        else {
            [void](Show-ThemedMessage -Owner $dlg -Title 'Browse' -Message 'Select a row, then click OK.' -Buttons OK -Icon Info)
        }
    })

    $btnCancel.Add_Click({ $script:DialogResult = $null; $dlg.Close() })

    $dlg.Add_Loaded({ $txtFilter.Focus() })
    [void]$dlg.ShowDialog()
    return $script:DialogResult
}

# =============================================================================
# Device collection browse: cached bulk load, then the shared folder-tree
# picker. The shared picker has no Refresh button, so -Force (Shift held on
# the Browse button) reloads. Show-CollectionBrowseItem returns the picked
# row (Name, CollectionID, ...); Show-CollectionBrowse returns its name.
# =============================================================================
function Show-CollectionBrowseItem {
    param(
        [Parameter(Mandatory)]$Owner,
        [switch]$Force
    )

    $loaded = Get-BrowseList -Owner $Owner -Title 'Browse device collections' -Type 'Collections' -Force:$Force
    if ($null -eq $loaded) { return $null }
    if (@($loaded.Items).Count -eq 0) {
        [void](Show-ThemedMessage -Owner $Owner -Title 'Browse' -Message 'No device collections were returned by the site.' -Buttons OK -Icon Info)
        return $null
    }

    return Show-CollectionPickerDialog -Owner $Owner -Collections @($loaded.Items) -Folders @($loaded.Folders) -Title 'Pick device collection'
}

function Show-CollectionBrowse {
    param(
        [Parameter(Mandatory)]$Owner,
        [switch]$Force
    )

    $picked = Show-CollectionBrowseItem -Owner $Owner -Force:$Force
    if ($picked) { return [string]$picked.Name }
    return $null
}

# =============================================================================
# Apps validation and deploy
# =============================================================================
$script:InvokeAppsValidate = {
    Reset-ValidationUi

    $appName = $txtTargetName.Text
    if ($null -ne $appName) { $appName = $appName.Trim() }
    $collName = $txtCollection.Text
    if ($null -ne $collName) { $collName = $collName.Trim() }

    if ([string]::IsNullOrWhiteSpace($appName) -or [string]::IsNullOrWhiteSpace($collName)) {
        Add-LogLine -Message 'Application and Collection are required.'
        Set-StatusText -Text 'Fill in Application and Collection to validate.'
        return
    }

    if (-not (Connect-IfNeeded)) { return }

    Set-StatusText -Text 'Validating...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        $app = Test-ApplicationExists -ApplicationName $appName
        if ($null -eq $app) {
            Set-CheckGlyph -Index 1 -State 'Fail'
            Set-StatusText -Text 'Application not found.'
            return
        }
        Set-CheckGlyph -Index 1 -State 'Pass'

        $dist = Test-ContentDistributed -Application $app
        if ($dist.IsFullyDistributed) {
            Set-CheckGlyph -Index 2 -State 'Pass'
        } else {
            Set-CheckGlyph -Index 2 -State 'Fail'
            Add-LogLine -Message ('Content not fully distributed: {0}/{1} DP(s) succeeded.' -f $dist.NumberSuccess, $dist.Targeted)
        }

        $col = Test-CollectionValid -CollectionName $collName
        if ($null -eq $col) {
            Set-CheckGlyph -Index 3 -State 'Fail'
            Set-StatusText -Text 'Collection invalid.'
            return
        }
        Set-CheckGlyph -Index 3 -State 'Pass'

        $safe = Test-CollectionSafe -Collection $col
        if ($safe.IsSafe) {
            Set-CheckGlyph -Index 4 -State 'Pass'
        } else {
            Set-CheckGlyph -Index 4 -State 'Fail'
            Add-LogLine -Message ('Collection blocked: {0}' -f $safe.Reason)
            Set-StatusText -Text 'Collection unsafe.'
            return
        }

        $dup = Test-DuplicateDeployment -ApplicationName $appName -CollectionName $collName
        $dupFailure = Get-DuplicateCheckFailure -Result $dup
        if ($dupFailure) {
            Set-CheckGlyph -Index 5 -State 'Fail'
            Add-LogLine -Message ('Duplicate check failed; deployment is blocked: {0}' -f $dupFailure.Error)
        } elseif ($null -eq $dup) {
            Set-CheckGlyph -Index 5 -State 'Pass'
        } else {
            Set-CheckGlyph -Index 5 -State 'Fail'
            Add-LogLine -Message 'Duplicate deployment already exists.'
        }

        $script:ValidatedApp        = $app
        $script:ValidatedCollection = $col

        $passAll = $true
        for ($i = 1; $i -le 5; $i++) {
            $g = switch ($i) { 1 {$glyphCheck1} 2 {$glyphCheck2} 3 {$glyphCheck3} 4 {$glyphCheck4} 5 {$glyphCheck5} }
            if ($g.Text -ne [string]$script:GlyphPass) { $passAll = $false; break }
        }
        $script:AllChecksPassed = $passAll

        if ($script:AllChecksPassed) {
            $btnDeploy.Visibility = [System.Windows.Visibility]::Visible
            Set-StatusText -Text 'All checks passed. Ready to deploy.'
        } else {
            Set-StatusText -Text 'Validation complete. Fix failing checks before deploying.'
        }
    }
    finally {
        $window.Cursor = $null
    }
}

function Update-PackageProgramDropdown {
    param($Package)
    $cboProgram.Items.Clear()
    if ($null -eq $Package) { return }
    $programs = @(Get-CMPackagePrograms -Package $Package)
    foreach ($p in $programs) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $p.ProgramName
        $item.Tag     = $p
        [void]$cboProgram.Items.Add($item)
    }
    if ($cboProgram.Items.Count -eq 1) {
        $cboProgram.SelectedIndex = 0
    }
    elseif ($cboProgram.Items.Count -gt 1) {
        Add-LogLine -Message ('Package has {0} programs. Select one.' -f $cboProgram.Items.Count)
    }
    else {
        Add-LogLine -Message 'Package has no programs defined.'
    }
}

# =============================================================================
# Packages validation and deploy
# =============================================================================
$script:InvokePackagesValidate = {
    Reset-ValidationUi

    $pkgName = $txtTargetName.Text
    if ($null -ne $pkgName) { $pkgName = $pkgName.Trim() }
    $collName = $txtCollection.Text
    if ($null -ne $collName) { $collName = $collName.Trim() }

    if ([string]::IsNullOrWhiteSpace($pkgName) -or [string]::IsNullOrWhiteSpace($collName)) {
        Add-LogLine -Message 'Package and Collection are required.'
        Set-StatusText -Text 'Fill in Package and Collection to validate.'
        return
    }

    if (-not (Connect-IfNeeded)) { return }

    Set-StatusText -Text 'Validating...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        # Check 1: package exists + a program is selected
        $pkg = Test-PackageExists -PackageName $pkgName
        if ($null -eq $pkg) {
            Set-CheckGlyph -Index 1 -State 'Fail'
            Set-StatusText -Text 'Package not found.'
            return
        }
        if ($cboProgram.Items.Count -eq 0) {
            Update-PackageProgramDropdown -Package $pkg
        }
        if ($null -eq $cboProgram.SelectedItem) {
            Set-CheckGlyph -Index 1 -State 'Fail'
            Add-LogLine -Message 'Select a program from the dropdown.'
            Set-StatusText -Text 'Program selection required.'
            return
        }
        $programInfo = $cboProgram.SelectedItem.Tag
        $programName = [string]$cboProgram.SelectedItem.Content
        Set-CheckGlyph -Index 1 -State 'Pass'

        # Check 2: content distributed (shared helper works for packages too -- it keys off PackageID)
        $dist = Test-ContentDistributed -Application $pkg
        if ($dist.IsFullyDistributed) {
            Set-CheckGlyph -Index 2 -State 'Pass'
        } else {
            Set-CheckGlyph -Index 2 -State 'Fail'
            Add-LogLine -Message ('Content not fully distributed: {0}/{1} DP(s) succeeded.' -f $dist.NumberSuccess, $dist.Targeted)
        }

        # Check 3: collection valid
        $col = Test-CollectionValid -CollectionName $collName
        if ($null -eq $col) {
            Set-CheckGlyph -Index 3 -State 'Fail'
            Set-StatusText -Text 'Collection invalid.'
            return
        }
        Set-CheckGlyph -Index 3 -State 'Pass'

        # Check 4: collection safe
        $safe = Test-CollectionSafe -Collection $col
        if ($safe.IsSafe) {
            Set-CheckGlyph -Index 4 -State 'Pass'
        } else {
            Set-CheckGlyph -Index 4 -State 'Fail'
            Add-LogLine -Message ('Collection blocked: {0}' -f $safe.Reason)
            Set-StatusText -Text 'Collection unsafe.'
            return
        }

        # Check 5: duplicate package+program+collection deployment
        $dup = Test-DuplicatePackageDeployment -PackageID $pkg.PackageID -ProgramName $programName -CollectionName $collName
        $dupFailure = Get-DuplicateCheckFailure -Result $dup
        if ($dupFailure) {
            Set-CheckGlyph -Index 5 -State 'Fail'
            Add-LogLine -Message ('Duplicate check failed; deployment is blocked: {0}' -f $dupFailure.Error)
        } elseif ($null -eq $dup) {
            Set-CheckGlyph -Index 5 -State 'Pass'
        } else {
            Set-CheckGlyph -Index 5 -State 'Fail'
            Add-LogLine -Message 'Duplicate package/program deployment already exists.'
        }

        $script:ValidatedPackage    = $pkg
        $script:ValidatedProgram    = $programName
        $script:ValidatedCollection = $col

        $passAll = $true
        for ($i = 1; $i -le 5; $i++) {
            $g = switch ($i) { 1 {$glyphCheck1} 2 {$glyphCheck2} 3 {$glyphCheck3} 4 {$glyphCheck4} 5 {$glyphCheck5} }
            if ($g.Text -ne [string]$script:GlyphPass) { $passAll = $false; break }
        }
        $script:AllChecksPassed = $passAll

        if ($script:AllChecksPassed) {
            $btnDeploy.Visibility = [System.Windows.Visibility]::Visible
            Set-StatusText -Text 'All checks passed. Ready to deploy.'
        } else {
            Set-StatusText -Text 'Validation complete. Fix failing checks before deploying.'
        }
    }
    finally {
        $window.Cursor = $null
    }
}

$script:InvokePackagesDeploy = {
    if (-not $script:AllChecksPassed -or $null -eq $script:ValidatedPackage -or $null -eq $script:ValidatedProgram -or $null -eq $script:ValidatedCollection) {
        Add-LogLine -Message 'Run Validate first.'
        return
    }

    $purpose   = if ($radRequired.IsChecked) { 'Required' } else { 'Available' }
    $timeBasis = if ($radUtc.IsChecked)      { 'Utc' }      else { 'LocalTime' }

    $available = if ($dtpAvailable.SelectedDateTime) { [datetime]$dtpAvailable.SelectedDateTime } else { Get-Date }
    $deadline  = $null
    if ($purpose -eq 'Required') {
        $deadline = if ($dtpDeadline.SelectedDateTime) { [datetime]$dtpDeadline.SelectedDateTime } else { (Get-Date).AddHours(24) }
    }

    $schedule = Test-DeploymentSchedule -Available $available -Deadline $deadline -Purpose $purpose -Now $(if ($timeBasis -eq 'Utc') { [datetime]::UtcNow } else { Get-Date })
    if (-not $schedule.Ok) {
        [void](Show-ThemedMessage -Owner $window -Title 'Deployment blocked' -Message $schedule.Reason -Buttons OK -Icon Error)
        Set-StatusText -Text 'Deployment blocked: invalid schedule.'
        return
    }
    if ($schedule.Level -eq 'Warn') {
        $proceed = Show-ThemedMessage -Owner $window -Title 'Check deployment schedule' -Message $schedule.Reason -Buttons YesNo -Icon Warn
        if ($proceed -ne 'Yes') {
            Add-LogLine -Message 'Deployment cancelled: schedule warning not confirmed.'
            return
        }
    }

    $confirmMsg = ("Deploy:`n  Package: {0}`n  Program: {1}`n  -> {2} ({3} members)`n  Purpose: {4}`n`nContinue?" -f
        $script:ValidatedPackage.Name,
        $script:ValidatedProgram,
        $script:ValidatedCollection.Name,
        $script:ValidatedCollection.MemberCount,
        $purpose)
    $confirm = Show-ThemedMessage -Owner $window -Title 'Confirm deployment' -Message $confirmMsg -Buttons OKCancel -Icon Question
    if ($confirm -ne 'OK') {
        Add-LogLine -Message 'Deployment cancelled by user.'
        return
    }

    Set-StatusText -Text 'Deploying...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        # Read from the Packages-scoped controls (pnlPackageOptions).
        # Apps-scoped chkOverrideSW/chkRebootOutSW/chkMetered live on
        # pnlRequiredOptions which is Apps-only, so those were silently
        # always $false for Packages. Fast/Slow network + Rerun were
        # never surfaced -- New-CMPackageDeployment was logging two
        # WARN lines ("-FastNetworkOption is not defined ...") on every
        # Packages deploy because defaults were used implicitly.
        $fastNet   = if ($cboPkgFastNetwork.SelectedItem) { [string]$cboPkgFastNetwork.SelectedItem.Tag } else { 'DownloadContentFromDistributionPointAndRunLocally' }
        $slowNet   = if ($cboPkgSlowNetwork.SelectedItem) { [string]$cboPkgSlowNetwork.SelectedItem.Tag } else { 'DoNotRunProgram' }
        $rerunBeh  = if ($cboPkgRerun.SelectedItem)       { [string]$cboPkgRerun.SelectedItem.Tag }       else { 'NeverRerunDeployedProgram' }
        $params = @{
            Package                    = $script:ValidatedPackage
            ProgramName                = $script:ValidatedProgram
            Collection                 = $script:ValidatedCollection
            DeployPurpose              = $purpose
            AvailableDateTime          = $available
            TimeBasedOn                = $timeBasis
            OverrideServiceWindow      = [bool]$chkPkgOverrideSW.IsChecked
            RebootOutsideServiceWindow = [bool]$chkPkgRebootOutSW.IsChecked
            UseMeteredNetwork          = [bool]$chkPkgMetered.IsChecked
            FastNetworkOption          = $fastNet
            SlowNetworkOption          = $slowNet
            RerunBehavior              = $rerunBeh
        }
        if ($deadline) { $params['DeadlineDateTime'] = $deadline }

        if (-not (Test-NoCurrentDuplicateDeployment -Type Package -TargetObject $script:ValidatedPackage `
                -Collection $script:ValidatedCollection -ProgramName $script:ValidatedProgram)) { return }
        if (-not (Test-DeploymentAuditReady)) { return }
        $result = Invoke-PackageDeployment @params

        $record = @{
            DeploymentType     = 'Package'
            ApplicationName    = $script:ValidatedPackage.Name
            ApplicationVersion = $script:ValidatedProgram
            CollectionName     = $script:ValidatedCollection.Name
            CollectionID       = $script:ValidatedCollection.CollectionID
            MemberCount        = $script:ValidatedCollection.MemberCount
            DeployPurpose      = $purpose
            DeadlineDateTime   = if ($deadline) { $deadline.ToString('yyyy-MM-ddTHH:mm:ss') } else { $null }
            DeploymentID       = $result.DeploymentID
            Result             = if ($result.Success) { 'Success' } else { ('Failed: {0}' -f $result.Error) }
        }
        $auditResult = Write-DeploymentLog -LogPath $script:DeploymentAuditLog -Record $record

        if ($result.Success) {
            Add-LogLine -Message ('Package deployment succeeded. DeploymentID={0}' -f $result.DeploymentID)
            Set-StatusText -Text 'Deployment succeeded.'
            $btnDeploy.Visibility = [System.Windows.Visibility]::Collapsed
            Reset-ValidationUi
        } else {
            Add-LogLine -Message ('Package deployment failed: {0}' -f $result.Error)
            Set-StatusText -Text 'Deployment failed.'
        }
        if (-not $auditResult.Success) { Show-AuditWriteFailure -AuditResult $auditResult -DeploymentCreated ([bool]$result.Success) -DeploymentID ([string]$result.DeploymentID) }
    }
    finally {
        $window.Cursor = $null
    }
}

# =============================================================================
# SUG validation and deploy
# =============================================================================
$script:InvokeSUGValidate = {
    Reset-ValidationUi

    $sugName = $txtTargetName.Text
    if ($null -ne $sugName) { $sugName = $sugName.Trim() }
    $collName = $txtCollection.Text
    if ($null -ne $collName) { $collName = $collName.Trim() }

    if ([string]::IsNullOrWhiteSpace($sugName) -or [string]::IsNullOrWhiteSpace($collName)) {
        Add-LogLine -Message 'Software Update Group and Collection are required.'
        Set-StatusText -Text 'Fill in Update Group and Collection to validate.'
        return
    }

    if (-not (Connect-IfNeeded)) { return }

    Set-StatusText -Text 'Validating...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        # Check 1: SUG exists
        $sug = Test-SUGExists -SUGName $sugName
        if ($null -eq $sug) {
            Set-CheckGlyph -Index 1 -State 'Fail'
            Set-StatusText -Text 'SUG not found.'
            return
        }
        Set-CheckGlyph -Index 1 -State 'Pass'

        # Check 2: "Content distributed" doesn't apply to SUGs the same way -- they
        # reference updates whose content lives in deployment packages. Mark
        # pass informationally; operator still has to ensure update packages are
        # distributed separately.
        Set-CheckGlyph -Index 2 -State 'Pass'

        # Check 3: collection valid
        $col = Test-CollectionValid -CollectionName $collName
        if ($null -eq $col) {
            Set-CheckGlyph -Index 3 -State 'Fail'
            Set-StatusText -Text 'Collection invalid.'
            return
        }
        Set-CheckGlyph -Index 3 -State 'Pass'

        # Check 4: collection safe
        $safe = Test-CollectionSafe -Collection $col
        if ($safe.IsSafe) {
            Set-CheckGlyph -Index 4 -State 'Pass'
        } else {
            Set-CheckGlyph -Index 4 -State 'Fail'
            Add-LogLine -Message ('Collection blocked: {0}' -f $safe.Reason)
            Set-StatusText -Text 'Collection unsafe.'
            return
        }

        # Check 5: duplicate SUG deployment
        $dup = Test-DuplicateSUGDeployment -SUGName $sug.LocalizedDisplayName -CollectionName $collName
        $dupFailure = Get-DuplicateCheckFailure -Result $dup
        if ($dupFailure) {
            Set-CheckGlyph -Index 5 -State 'Fail'
            Add-LogLine -Message ('Duplicate check failed; deployment is blocked: {0}' -f $dupFailure.Error)
        } elseif ($null -eq $dup) {
            Set-CheckGlyph -Index 5 -State 'Pass'
        } else {
            Set-CheckGlyph -Index 5 -State 'Fail'
            Add-LogLine -Message 'Duplicate SUG deployment already exists.'
        }

        $script:ValidatedSUG        = $sug
        $script:ValidatedCollection = $col

        $passAll = $true
        for ($i = 1; $i -le 5; $i++) {
            $g = switch ($i) { 1 {$glyphCheck1} 2 {$glyphCheck2} 3 {$glyphCheck3} 4 {$glyphCheck4} 5 {$glyphCheck5} }
            if ($g.Text -ne [string]$script:GlyphPass) { $passAll = $false; break }
        }
        $script:AllChecksPassed = $passAll

        if ($script:AllChecksPassed) {
            $btnDeploy.Visibility = [System.Windows.Visibility]::Visible
            Set-StatusText -Text 'All checks passed. Ready to deploy.'
        } else {
            Set-StatusText -Text 'Validation complete. Fix failing checks before deploying.'
        }
    }
    finally {
        $window.Cursor = $null
    }
}

$script:InvokeSUGDeploy = {
    if (-not $script:AllChecksPassed -or $null -eq $script:ValidatedSUG -or $null -eq $script:ValidatedCollection) {
        Add-LogLine -Message 'Run Validate first.'
        return
    }

    $purpose   = if ($radRequired.IsChecked) { 'Required' } else { 'Available' }
    $timeBasis = if ($radUtc.IsChecked)      { 'Utc' }      else { 'LocalTime' }
    $notifTag  = if ($cboNotification.SelectedItem) { [string]$cboNotification.SelectedItem.Tag } else { 'DisplayAll' }

    $available = if ($dtpAvailable.SelectedDateTime) { [datetime]$dtpAvailable.SelectedDateTime } else { Get-Date }
    $deadline  = $null
    if ($purpose -eq 'Required') {
        $deadline = if ($dtpDeadline.SelectedDateTime) { [datetime]$dtpDeadline.SelectedDateTime } else { (Get-Date).AddHours(24) }
    }

    $schedule = Test-DeploymentSchedule -Available $available -Deadline $deadline -Purpose $purpose -Now $(if ($timeBasis -eq 'Utc') { [datetime]::UtcNow } else { Get-Date })
    if (-not $schedule.Ok) {
        [void](Show-ThemedMessage -Owner $window -Title 'Deployment blocked' -Message $schedule.Reason -Buttons OK -Icon Error)
        Set-StatusText -Text 'Deployment blocked: invalid schedule.'
        return
    }
    if ($schedule.Level -eq 'Warn') {
        $proceed = Show-ThemedMessage -Owner $window -Title 'Check deployment schedule' -Message $schedule.Reason -Buttons YesNo -Icon Warn
        if ($proceed -ne 'Yes') {
            Add-LogLine -Message 'Deployment cancelled: schedule warning not confirmed.'
            return
        }
    }

    $notification = Test-DeploymentNotification -Purpose $purpose -Notification $notifTag
    if (-not $notification.Ok) {
        [void](Show-ThemedMessage -Owner $window -Title 'Deployment blocked' -Message $notification.Reason -Buttons OK -Icon Error)
        Set-StatusText -Text 'Deployment blocked: invalid notification/purpose combo.'
        return
    }

    $confirmMsg = ("Deploy:`n  SUG: {0} ({1} updates)`n  -> {2} ({3} members)`n  Purpose: {4}`n`nContinue?" -f
        $script:ValidatedSUG.LocalizedDisplayName,
        $script:ValidatedSUG.NumberOfUpdates,
        $script:ValidatedCollection.Name,
        $script:ValidatedCollection.MemberCount,
        $purpose)
    $confirm = Show-ThemedMessage -Owner $window -Title 'Confirm deployment' -Message $confirmMsg -Buttons OKCancel -Icon Question
    if ($confirm -ne 'OK') {
        Add-LogLine -Message 'Deployment cancelled by user.'
        return
    }

    Set-StatusText -Text 'Deploying...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        # Read from the SUG-scoped checkboxes (pnlSUGOptions). The
        # apps-scoped chkOverrideSW / chkRebootOutSW / chkMetered live
        # on pnlRequiredOptions which is collapsed whenever Type=SUG,
        # so their IsChecked was always the default -- the SUG deploy
        # was silently ignoring these three bools for every run.
        $params = @{
            SUG                          = $script:ValidatedSUG
            Collection                   = $script:ValidatedCollection
            DeployPurpose                = $purpose
            AvailableDateTime            = $available
            TimeBasedOn                  = $timeBasis
            UserNotification             = $notifTag
            SoftwareInstallation         = [bool]$chkSugOverrideSW.IsChecked
            AllowRestart                 = [bool]$chkSugAllowRestart.IsChecked
            UseMeteredNetwork            = [bool]$chkSugMetered.IsChecked
            AllowBoundaryFallback        = [bool]$chkBoundaryFallback.IsChecked
            DownloadFromMicrosoftUpdate  = [bool]$chkMSFallback.IsChecked
            RequirePostRebootFullScan    = [bool]$chkFullScan.IsChecked
        }
        if ($deadline) { $params['DeadlineDateTime'] = $deadline }

        if (-not (Test-NoCurrentDuplicateDeployment -Type SUG -TargetObject $script:ValidatedSUG `
                -Collection $script:ValidatedCollection)) { return }
        if (-not (Test-DeploymentAuditReady)) { return }
        $result = Invoke-SUGDeployment @params

        $preview = Get-DeploymentPreview -TargetObject $script:ValidatedSUG -Collection $script:ValidatedCollection -DeploymentType 'SUG'
        $record = @{
            DeploymentType     = 'SUG'
            ApplicationName    = $preview.ApplicationName
            ApplicationVersion = $preview.ApplicationVersion
            CollectionName     = $preview.CollectionName
            CollectionID       = $preview.CollectionID
            MemberCount        = $preview.MemberCount
            DeployPurpose      = $purpose
            DeadlineDateTime   = if ($deadline) { $deadline.ToString('yyyy-MM-ddTHH:mm:ss') } else { $null }
            DeploymentID       = $result.DeploymentID
            Result             = if ($result.Success) { 'Success' } else { ('Failed: {0}' -f $result.Error) }
        }
        $auditResult = Write-DeploymentLog -LogPath $script:DeploymentAuditLog -Record $record

        if ($result.Success) {
            Add-LogLine -Message ('SUG deployment succeeded. AssignmentID={0}' -f $result.DeploymentID)
            Set-StatusText -Text 'Deployment succeeded.'
            $btnDeploy.Visibility = [System.Windows.Visibility]::Collapsed
            Reset-ValidationUi
        } else {
            Add-LogLine -Message ('SUG deployment failed: {0}' -f $result.Error)
            Set-StatusText -Text 'Deployment failed.'
        }
        if (-not $auditResult.Success) { Show-AuditWriteFailure -AuditResult $auditResult -DeploymentCreated ([bool]$result.Success) -DeploymentID ([string]$result.DeploymentID) }
    }
    finally {
        $window.Cursor = $null
    }
}

# =============================================================================
# Task Sequences validation and deploy
# =============================================================================
$script:InvokeTaskSequencesValidate = {
    Reset-ValidationUi

    $tsName = $txtTargetName.Text
    if ($null -ne $tsName) { $tsName = $tsName.Trim() }
    $collName = $txtCollection.Text
    if ($null -ne $collName) { $collName = $collName.Trim() }

    if ([string]::IsNullOrWhiteSpace($tsName) -or [string]::IsNullOrWhiteSpace($collName)) {
        Add-LogLine -Message 'Task Sequence and Collection are required.'
        Set-StatusText -Text 'Fill in Task Sequence and Collection to validate.'
        return
    }

    if (-not (Connect-IfNeeded)) { return }

    Set-StatusText -Text 'Validating...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        # Check 1: TS exists
        $ts = Test-TaskSequenceExists -TaskSequenceName $tsName
        if ($null -eq $ts) {
            Set-CheckGlyph -Index 1 -State 'Fail'
            Set-StatusText -Text 'Task sequence not found.'
            return
        }
        Set-CheckGlyph -Index 1 -State 'Pass'

        # Check 2: content distribution -- N/A for TS. A task-sequence
        # "package" is metadata only; all distributable content lives
        # in the TS's referenced boot images / packages / apps, each
        # tracked independently. Walking that reference tree is a v1.1
        # candidate (Get-CMTaskSequenceDeployment's -InputObject
        # Reference collection). For v1.0, mark PASS with an info log;
        # ConfigMgr's deploy cmdlet will still error if a referenced item's
        # content isn't on a reachable DP. User confirmed "content is
        # fully distributed" surface via the console for this case.
        Set-CheckGlyph -Index 2 -State 'Pass'
        Add-LogLine -Message 'Check 2 (content) skipped for TS: referenced-content distribution enforced by ConfigMgr at deploy time. Verify in console if unsure.'

        # Check 3: collection valid
        $col = Test-CollectionValid -CollectionName $collName
        if ($null -eq $col) {
            Set-CheckGlyph -Index 3 -State 'Fail'
            Set-StatusText -Text 'Collection invalid.'
            return
        }
        Set-CheckGlyph -Index 3 -State 'Pass'

        # Check 4: collection safe
        $safe = Test-CollectionSafe -Collection $col
        if ($safe.IsSafe) {
            Set-CheckGlyph -Index 4 -State 'Pass'
        } else {
            Set-CheckGlyph -Index 4 -State 'Fail'
            Add-LogLine -Message ('Collection blocked: {0}' -f $safe.Reason)
            Set-StatusText -Text 'Collection unsafe.'
            return
        }

        # Check 5: duplicate TS deployment
        $dup = Test-DuplicateTaskSequenceDeployment -TaskSequencePackageId $ts.PackageID -CollectionName $collName
        $dupFailure = Get-DuplicateCheckFailure -Result $dup
        if ($dupFailure) {
            Set-CheckGlyph -Index 5 -State 'Fail'
            Add-LogLine -Message ('Duplicate check failed; deployment is blocked: {0}' -f $dupFailure.Error)
        } elseif ($null -eq $dup) {
            Set-CheckGlyph -Index 5 -State 'Pass'
        } else {
            Set-CheckGlyph -Index 5 -State 'Fail'
            Add-LogLine -Message 'Duplicate task sequence deployment already exists.'
        }

        $script:ValidatedTaskSequence = $ts
        $script:ValidatedCollection   = $col

        $passAll = $true
        for ($i = 1; $i -le 5; $i++) {
            $g = switch ($i) { 1 {$glyphCheck1} 2 {$glyphCheck2} 3 {$glyphCheck3} 4 {$glyphCheck4} 5 {$glyphCheck5} }
            if ($g.Text -ne [string]$script:GlyphPass) { $passAll = $false; break }
        }
        $script:AllChecksPassed = $passAll

        if ($script:AllChecksPassed) {
            $btnDeploy.Visibility = [System.Windows.Visibility]::Visible
            Set-StatusText -Text 'All checks passed. Ready to deploy.'
        } else {
            Set-StatusText -Text 'Validation complete. Fix failing checks before deploying.'
        }
    }
    finally {
        $window.Cursor = $null
    }
}

$script:InvokeTaskSequencesDeploy = {
    if (-not $script:AllChecksPassed -or $null -eq $script:ValidatedTaskSequence -or $null -eq $script:ValidatedCollection) {
        Add-LogLine -Message 'Run Validate first.'
        return
    }

    $purpose   = if ($radRequired.IsChecked) { 'Required' } else { 'Available' }
    $timeBasis = if ($radUtc.IsChecked)      { 'Utc' }      else { 'LocalTime' }
    $avail     = if ($cboNotification.SelectedItem) { [string]$cboNotification.SelectedItem.Tag } else { 'Clients' }

    $available = if ($dtpAvailable.SelectedDateTime) { [datetime]$dtpAvailable.SelectedDateTime } else { Get-Date }
    $deadline  = $null
    if ($purpose -eq 'Required') {
        $deadline = if ($dtpDeadline.SelectedDateTime) { [datetime]$dtpDeadline.SelectedDateTime } else { (Get-Date).AddHours(24) }
    }

    $schedule = Test-DeploymentSchedule -Available $available -Deadline $deadline -Purpose $purpose -Now $(if ($timeBasis -eq 'Utc') { [datetime]::UtcNow } else { Get-Date })
    if (-not $schedule.Ok) {
        [void](Show-ThemedMessage -Owner $window -Title 'Deployment blocked' -Message $schedule.Reason -Buttons OK -Icon Error)
        Set-StatusText -Text 'Deployment blocked: invalid schedule.'
        return
    }
    if ($schedule.Level -eq 'Warn') {
        $proceed = Show-ThemedMessage -Owner $window -Title 'Check deployment schedule' -Message $schedule.Reason -Buttons YesNo -Icon Warn
        if ($proceed -ne 'Yes') {
            Add-LogLine -Message 'Deployment cancelled: schedule warning not confirmed.'
            return
        }
    }

    $confirmMsg = ("Deploy:`n  Task Sequence: {0}`n  -> {1} ({2} members)`n  Purpose: {3}`n  Availability: {4}`n`nContinue?" -f
        $script:ValidatedTaskSequence.Name,
        $script:ValidatedCollection.Name,
        $script:ValidatedCollection.MemberCount,
        $purpose,
        $avail)
    $confirm = Show-ThemedMessage -Owner $window -Title 'Confirm deployment' -Message $confirmMsg -Buttons OKCancel -Icon Question
    if ($confirm -ne 'OK') {
        Add-LogLine -Message 'Deployment cancelled by user.'
        return
    }

    Set-StatusText -Text 'Deploying...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        # Same refactor as SUG: read from the TS-scoped checkboxes
        # (pnlTaskSequenceOptions). The apps-scoped trio lives on
        # pnlRequiredOptions which is collapsed whenever Type=TS, so
        # reading chkOverrideSW/chkRebootOutSW/chkMetered here was
        # silently always $false.
        $params = @{
            TaskSequence               = $script:ValidatedTaskSequence
            Collection                 = $script:ValidatedCollection
            DeployPurpose              = $purpose
            AvailableDateTime          = $available
            Availability               = $avail
            TimeBasedOn                = $timeBasis
            ShowTaskSequenceProgress   = [bool]$chkShowProgress.IsChecked
            OverrideServiceWindow      = [bool]$chkTsOverrideSW.IsChecked
            RebootOutsideServiceWindow = [bool]$chkTsRebootOutSW.IsChecked
            UseMeteredNetwork          = [bool]$chkTsMetered.IsChecked
        }
        if ($deadline) { $params['DeadlineDateTime'] = $deadline }

        if (-not (Test-NoCurrentDuplicateDeployment -Type TaskSequence -TargetObject $script:ValidatedTaskSequence `
                -Collection $script:ValidatedCollection)) { return }
        if (-not (Test-DeploymentAuditReady)) { return }
        $result = Invoke-TaskSequenceDeployment @params

        $record = @{
            DeploymentType     = 'TaskSequence'
            ApplicationName    = $script:ValidatedTaskSequence.Name
            ApplicationVersion = $avail
            CollectionName     = $script:ValidatedCollection.Name
            CollectionID       = $script:ValidatedCollection.CollectionID
            MemberCount        = $script:ValidatedCollection.MemberCount
            DeployPurpose      = $purpose
            DeadlineDateTime   = if ($deadline) { $deadline.ToString('yyyy-MM-ddTHH:mm:ss') } else { $null }
            DeploymentID       = $result.DeploymentID
            Result             = if ($result.Success) { 'Success' } else { ('Failed: {0}' -f $result.Error) }
        }
        $auditResult = Write-DeploymentLog -LogPath $script:DeploymentAuditLog -Record $record

        if ($result.Success) {
            Add-LogLine -Message ('Task sequence deployment succeeded. DeploymentID={0}' -f $result.DeploymentID)
            Set-StatusText -Text 'Deployment succeeded.'
            $btnDeploy.Visibility = [System.Windows.Visibility]::Collapsed
            Reset-ValidationUi
        } else {
            Add-LogLine -Message ('Task sequence deployment failed: {0}' -f $result.Error)
            Set-StatusText -Text 'Deployment failed.'
        }
        if (-not $auditResult.Success) { Show-AuditWriteFailure -AuditResult $auditResult -DeploymentCreated ([bool]$result.Success) -DeploymentID ([string]$result.DeploymentID) }
    }
    finally {
        $window.Cursor = $null
    }
}

$script:InvokeAppsDeploy = {
    if (-not $script:AllChecksPassed -or $null -eq $script:ValidatedApp -or $null -eq $script:ValidatedCollection) {
        Add-LogLine -Message 'Run Validate first.'
        return
    }

    $purpose = if ($radRequired.IsChecked) { 'Required' } else { 'Available' }
    $timeBasis = if ($radUtc.IsChecked) { 'Utc' } else { 'LocalTime' }

    $notifTag = if ($cboNotification.SelectedItem) {
        [string]$cboNotification.SelectedItem.Tag
    } else {
        'DisplayAll'
    }

    $available = if ($dtpAvailable.SelectedDateTime) { [datetime]$dtpAvailable.SelectedDateTime } else { Get-Date }
    $deadline  = $null
    if ($purpose -eq 'Required') {
        $deadline = if ($dtpDeadline.SelectedDateTime) { [datetime]$dtpDeadline.SelectedDateTime } else { (Get-Date).AddHours(24) }
    }

    # Schedule sanity. Block inverted Required deadlines; warn-and-confirm
    # on backdated Available. Stops accidental "fires immediately on every
    # client" moments before the ConfigMgr cmdlet accepts them.
    $schedule = Test-DeploymentSchedule -Available $available -Deadline $deadline -Purpose $purpose -Now $(if ($timeBasis -eq 'Utc') { [datetime]::UtcNow } else { Get-Date })
    if (-not $schedule.Ok) {
        [void](Show-ThemedMessage -Owner $window -Title 'Deployment blocked' -Message $schedule.Reason -Buttons OK -Icon Error)
        Set-StatusText -Text 'Deployment blocked: invalid schedule.'
        return
    }
    if ($schedule.Level -eq 'Warn') {
        $proceed = Show-ThemedMessage -Owner $window -Title 'Check deployment schedule' -Message $schedule.Reason -Buttons YesNo -Icon Warn
        if ($proceed -ne 'Yes') {
            Add-LogLine -Message 'Deployment cancelled: schedule warning not confirmed.'
            return
        }
    }

    $notification = Test-DeploymentNotification -Purpose $purpose -Notification $notifTag
    if (-not $notification.Ok) {
        [void](Show-ThemedMessage -Owner $window -Title 'Deployment blocked' -Message $notification.Reason -Buttons OK -Icon Error)
        Set-StatusText -Text 'Deployment blocked: invalid notification/purpose combo.'
        return
    }

    $confirmMsg = ("Deploy:`n  {0} v{1}`n  -> {2} ({3} members)`n  Purpose: {4}`n`nContinue?" -f
        $script:ValidatedApp.LocalizedDisplayName,
        $script:ValidatedApp.SoftwareVersion,
        $script:ValidatedCollection.Name,
        $script:ValidatedCollection.MemberCount,
        $purpose)
    $confirm = Show-ThemedMessage -Owner $window -Title 'Confirm deployment' -Message $confirmMsg -Buttons OKCancel -Icon Question
    if ($confirm -ne 'OK') {
        Add-LogLine -Message 'Deployment cancelled by user.'
        return
    }

    Set-StatusText -Text 'Deploying...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        $params = @{
            Application                = $script:ValidatedApp
            Collection                 = $script:ValidatedCollection
            DeployPurpose              = $purpose
            AvailableDateTime          = $available
            TimeBasedOn                = $timeBasis
            UserNotification           = $notifTag
            OverrideServiceWindow      = [bool]$chkOverrideSW.IsChecked
            RebootOutsideServiceWindow = [bool]$chkRebootOutSW.IsChecked
            UseMeteredNetwork          = [bool]$chkMetered.IsChecked
        }
        if ($deadline) { $params['DeadlineDateTime'] = $deadline }

        if (-not (Test-NoCurrentDuplicateDeployment -Type Application -TargetObject $script:ValidatedApp `
                -Collection $script:ValidatedCollection)) { return }
        if (-not (Test-DeploymentAuditReady)) { return }
        $result = Invoke-ApplicationDeployment @params

        $preview = Get-DeploymentPreview -TargetObject $script:ValidatedApp -Collection $script:ValidatedCollection -DeploymentType 'Application'
        $record = @{
            DeploymentType     = 'Application'
            ApplicationName    = $preview.ApplicationName
            ApplicationVersion = $preview.ApplicationVersion
            CollectionName     = $preview.CollectionName
            CollectionID       = $preview.CollectionID
            MemberCount        = $preview.MemberCount
            DeployPurpose      = $purpose
            DeadlineDateTime   = if ($deadline) { $deadline.ToString('yyyy-MM-ddTHH:mm:ss') } else { $null }
            DeploymentID       = $result.DeploymentID
            Result             = if ($result.Success) { 'Success' } else { ('Failed: {0}' -f $result.Error) }
        }
        $auditResult = Write-DeploymentLog -LogPath $script:DeploymentAuditLog -Record $record

        if ($result.Success) {
            Add-LogLine -Message ('Deployment succeeded. AssignmentID={0}' -f $result.DeploymentID)
            Set-StatusText -Text 'Deployment succeeded.'
            $btnDeploy.Visibility = [System.Windows.Visibility]::Collapsed
            Reset-ValidationUi
        } else {
            Add-LogLine -Message ('Deployment failed: {0}' -f $result.Error)
            Set-StatusText -Text 'Deployment failed.'
        }
        if (-not $auditResult.Success) { Show-AuditWriteFailure -AuditResult $auditResult -DeploymentCreated ([bool]$result.Success) -DeploymentID ([string]$result.DeploymentID) }
    }
    finally {
        $window.Cursor = $null
    }
}

# =============================================================================
# DP groups: cache + picker + distribute
# =============================================================================
function Get-DPGroupsLoaded {
    param([switch]$ForceReload)
    if ($ForceReload) { $global:DPGroupsCache = $null }
    if ($null -ne $global:DPGroupsCache) { return $global:DPGroupsCache }
    if (-not (Connect-IfNeeded)) { return @() }
    $names = @(Get-DPGroupList | ForEach-Object { [string]$_.Name })
    $global:DPGroupsCache = $names
    Add-LogLine -Message ('Loaded {0} DP group(s).' -f $names.Count)
    return $names
}

function Update-DPPickerButtonLabel {
    $sel = @($global:SelectedDPGroups)
    $all = @($global:DPGroupsCache)
    if ($sel.Count -eq 0) {
        $btnDPPicker.Content = 'Select DP groups...'
    }
    elseif ($all.Count -gt 0 -and $sel.Count -eq $all.Count) {
        $btnDPPicker.Content = ('All {0} DP groups selected' -f $all.Count)
    }
    else {
        $btnDPPicker.Content = ('{0} of {1} DP groups selected' -f $sel.Count, [math]::Max($all.Count, $sel.Count))
    }
}

function Get-CurrentDistributionTarget {
    # Look up the target object by what's currently in the textbox,
    # so the DP picker can pre-check already-targeted groups even if
    # the user hasn't clicked Validate yet.
    $name = $txtTargetName.Text
    if ($null -ne $name) { $name = $name.Trim() }
    if ([string]::IsNullOrWhiteSpace($name)) { return $null }
    switch ($script:CurrentType) {
        'Apps'          { return Test-ApplicationExists -ApplicationName $name }
        'Packages'      { return Test-PackageExists     -PackageName     $name }
        'TaskSequences' { return Test-TaskSequenceExists -TaskSequenceName $name }
        default         { return $null }
    }
}

function Get-CurrentDistributionType {
    switch ($script:CurrentType) {
        'Apps'          { return 'Application' }
        'Packages'      { return 'Package' }
        'TaskSequences' { return 'TaskSequence' }
        default         { return $null }
    }
}

function Show-DPGroupPickerDialog {
    param(
        [Parameter(Mandatory)]$Owner,
        [Parameter(Mandatory)][string[]]$AllGroupNames,
        [string[]]$PreSelected       = @(),
        [string[]]$AlreadyTargeted   = @()
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Select DP groups"
    Width="520" Height="480"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="14">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Check the distribution point groups to distribute content to. Groups that already have this content are pre-checked." FontSize="12" TextWrapping="Wrap" Margin="0,0,0,10"/>

        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <Button x:Name="btnSelectAll"  Content="Select all"  MinWidth="100" Height="26" Margin="0,0,8,0"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnSelectNone" Content="Select none" MinWidth="100" Height="26"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>

        <ListBox x:Name="lstGroups" Grid.Row="2" BorderThickness="1" BorderBrush="{DynamicResource MahApps.Brushes.Gray8}">
            <ListBox.ItemTemplate>
                <DataTemplate>
                    <StackPanel Orientation="Horizontal">
                        <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" VerticalAlignment="Center"/>
                        <TextBlock Text="{Binding Name}" FontSize="12" Margin="8,0,0,0" VerticalAlignment="Center"/>
                        <TextBlock Text="{Binding Badge}" FontSize="11" Margin="8,0,0,0" VerticalAlignment="Center"
                                   Foreground="{DynamicResource MahApps.Brushes.Gray2}"/>
                    </StackPanel>
                </DataTemplate>
            </ListBox.ItemTemplate>
        </ListBox>

        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnOK"     Content="OK"     MinWidth="90" Height="30" Margin="0,0,8,0" IsDefault="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            <Button x:Name="btnCancel" Content="Cancel" MinWidth="90" Height="30" IsCancel="True"
                    Style="{DynamicResource MahApps.Styles.Button.Square}"
                    Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$xml = $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $dlg    = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg

    $theme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
    if ($theme) { [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $theme) }
    $dlg.Owner = $Owner
    try {
        $dlg.WindowTitleBrush          = $Owner.WindowTitleBrush
        $dlg.NonActiveWindowTitleBrush = $Owner.WindowTitleBrush
        $dlg.GlowBrush                 = $Owner.GlowBrush
        $dlg.NonActiveGlowBrush        = $Owner.GlowBrush
    } catch { }

    $btnSelectAll  = $dlg.FindName('btnSelectAll')
    $btnSelectNone = $dlg.FindName('btnSelectNone')
    $lstGroups     = $dlg.FindName('lstGroups')
    $btnOK         = $dlg.FindName('btnOK')
    $btnCancel     = $dlg.FindName('btnCancel')

    # Build row view models. Use an ObservableCollection so CheckBox TwoWay bindings
    # round-trip cleanly; each row carries Name/Selected/Badge.
    $rows = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
    foreach ($name in $AllGroupNames) {
        $preChecked = ($name -in $PreSelected) -or ($name -in $AlreadyTargeted)
        $badge = if ($name -in $AlreadyTargeted) { '(already targeted)' } else { '' }
        $row = [PSCustomObject]@{
            Name     = $name
            Selected = [bool]$preChecked
            Badge    = $badge
        }
        $rows.Add($row) | Out-Null
    }
    $lstGroups.ItemsSource = $rows

    # No .GetNewClosure() on any handler in this modal dialog -- same
    # rule as Show-ThemedMessage / Show-BrowseDialog: ShowDialog() is
    # still blocking this function when the handlers fire, so lexical
    # parent scope reaches $rows / $lstGroups / $dlg naturally, AND
    # $script:DPPickerResult writes land on the real factory-scope var
    # rather than a stripped copy. GetNewClosure on btnOK was causing
    # the picker to silently return $null no matter what the user
    # checked, which is why "dp no populate after me select" reproduced.
    $btnSelectAll.Add_Click({
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $rows[$i] | Add-Member -MemberType NoteProperty -Name 'Selected' -Value $true -Force
        }
        # Force refresh by re-binding
        $lstGroups.Items.Refresh()
    })

    $btnSelectNone.Add_Click({
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $rows[$i] | Add-Member -MemberType NoteProperty -Name 'Selected' -Value $false -Force
        }
        $lstGroups.Items.Refresh()
    })

    $script:DPPickerResult = $null
    $btnOK.Add_Click({
        $selected = @($rows | Where-Object { $_.Selected } | ForEach-Object { [string]$_.Name })
        $script:DPPickerResult = $selected
        $dlg.Close()
    })

    $btnCancel.Add_Click({ $script:DPPickerResult = $null; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:DPPickerResult
}

# =============================================================================
# Options dialog (Connection, Logging, History, Templates, About)
# =============================================================================
function New-ConnectionPanel {
    $grid = New-Object System.Windows.Controls.StackPanel
    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = 'Configuration Manager Connection'
    $hdr.FontSize = 18
    $hdr.FontWeight = 'SemiBold'
    $hdr.Margin = '0,0,0,12'
    [void]$grid.Children.Add($hdr)

    $g = New-Object System.Windows.Controls.Grid
    2..0 | ForEach-Object {
        $rd = New-Object System.Windows.Controls.RowDefinition
        $rd.Height = 'Auto'
        [void]$g.RowDefinitions.Add($rd)
    }
    $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = '120'
    $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = '*'
    [void]$g.ColumnDefinitions.Add($c1)
    [void]$g.ColumnDefinitions.Add($c2)

    $lblSite = New-Object System.Windows.Controls.TextBlock
    $lblSite.Text = 'Site Code:'
    $lblSite.FontSize = 12
    $lblSite.VerticalAlignment = 'Center'
    $lblSite.Margin = '0,4,8,4'
    [System.Windows.Controls.Grid]::SetRow($lblSite, 0)
    [System.Windows.Controls.Grid]::SetColumn($lblSite, 0)
    [void]$g.Children.Add($lblSite)

    $txtSite = New-Object System.Windows.Controls.TextBox
    $txtSite.Name = 'txtSite'
    $txtSite.Text = [string]$global:Prefs['SiteCode']
    $txtSite.FontSize = 12
    $txtSite.Height = 28
    $txtSite.VerticalContentAlignment = 'Center'
    $txtSite.MaxLength = 3
    $txtSite.Margin = '0,4,0,4'
    $txtSite.Width = 80
    $txtSite.HorizontalAlignment = 'Left'
    [System.Windows.Controls.Grid]::SetRow($txtSite, 0)
    [System.Windows.Controls.Grid]::SetColumn($txtSite, 1)
    [void]$g.Children.Add($txtSite)

    $lblProv = New-Object System.Windows.Controls.TextBlock
    $lblProv.Text = 'SMS Provider:'
    $lblProv.FontSize = 12
    $lblProv.VerticalAlignment = 'Center'
    $lblProv.Margin = '0,4,8,4'
    [System.Windows.Controls.Grid]::SetRow($lblProv, 1)
    [System.Windows.Controls.Grid]::SetColumn($lblProv, 0)
    [void]$g.Children.Add($lblProv)

    $txtProv = New-Object System.Windows.Controls.TextBox
    $txtProv.Name = 'txtProv'
    $txtProv.Text = [string]$global:Prefs['SMSProvider']
    $txtProv.FontSize = 12
    $txtProv.Height = 28
    $txtProv.VerticalContentAlignment = 'Center'
    $txtProv.Margin = '0,4,0,4'
    [MahApps.Metro.Controls.TextBoxHelper]::SetWatermark($txtProv, 'server.fqdn')
    [System.Windows.Controls.Grid]::SetRow($txtProv, 1)
    [System.Windows.Controls.Grid]::SetColumn($txtProv, 1)
    [void]$g.Children.Add($txtProv)

    [void]$grid.Children.Add($g)

    $note = New-Object System.Windows.Controls.TextBlock
    $note.Text = 'Connect reattempts on OK if the site code or provider changed.'
    $note.FontSize = 11
    # Muted-note Foreground pair per theme -- #808080 fails AA on both.
    $note.Foreground = if ($toggleTheme.IsOn) { $script:LogLabelDark } else { $script:LogLabelLight }
    $note.TextWrapping = 'Wrap'
    $note.Margin = '0,10,0,16'
    [void]$grid.Children.Add($note)

    $dpHdr = New-Object System.Windows.Controls.TextBlock
    $dpHdr.Text = 'Distribution point groups'
    $dpHdr.FontSize = 13
    $dpHdr.FontWeight = 'SemiBold'
    $dpHdr.Margin = '0,6,0,6'
    [void]$grid.Children.Add($dpHdr)

    $dpStatus = New-Object System.Windows.Controls.TextBlock
    $dpStatus.Name = 'lblDPStatus'
    $dpStatus.FontSize = 12
    $dpStatus.Margin = '0,0,0,8'
    $cached = if ($null -ne $global:DPGroupsCache) { @($global:DPGroupsCache).Count } else { 0 }
    $dpStatus.Text = if ($null -ne $global:DPGroupsCache) { ('Cached: {0} group(s).' -f $cached) } else { 'Not loaded yet. Will load on first Select.' }
    [void]$grid.Children.Add($dpStatus)

    $dpBtn = New-Object System.Windows.Controls.Button
    $dpBtn.Name = 'btnReloadDPGroups'
    $dpBtn.Content = 'Reload DP groups'
    $dpBtn.MinWidth = 160
    $dpBtn.Height = 28
    $dpBtn.HorizontalAlignment = 'Left'
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($dpBtn, [System.Windows.Controls.CharacterCasing]::Normal)
    $dpBtn.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'MahApps.Styles.Button.Square')
    [void]$grid.Children.Add($dpBtn)

    # Helper-indirection pattern (see New-TemplatesPanel for the mechanics).
    # Panel factory returns, so handlers fire after the factory is off the
    # stack. GetNewClosure on the outer handler is REQUIRED to capture
    # factory locals ($updateDpStatus, $dpStatus, etc.) but it strips
    # script-scope function lookup AND $script: writes. We wrap each
    # script-scope callee in a PLAIN local scriptblock; plain scriptblocks
    # carry the main-script SessionState at creation, so invoking them via
    # '&' resolves script functions and lands $script: writes correctly.
    $connectIfNeededCall  = { Connect-IfNeeded }
    $dpGroupsReloadCall   = { Get-DPGroupsLoaded -ForceReload }
    $updateDpPickerCall   = { Update-DPPickerButtonLabel }
    $updateDpStatus = {
        param($groups)
        $dpStatus.Text = ('Cached: {0} group(s).' -f @($groups).Count)
    }.GetNewClosure()
    $dpBtn.Add_Click({
        if (-not (& $connectIfNeededCall)) { return }
        $groups = & $dpGroupsReloadCall
        & $updateDpStatus $groups
        # If the user had selections that no longer exist after reload, prune them
        if (@($global:SelectedDPGroups).Count -gt 0) {
            $global:SelectedDPGroups = @($global:SelectedDPGroups | Where-Object { $_ -in $groups })
            & $updateDpPickerCall
        }
    }.GetNewClosure())

    # Commit runs AFTER the factory returns (master OK handler invokes
    # $panels[i].Commit later). GetNewClosure is required to reach $txtSite
    # / $txtProv. Script-scope effects (Disconnect-CMSite, Add-LogLine,
    # $script:ConnectedToCM write) are routed through a plain local
    # scriptblock per PS51-WPF-006 so they resolve against main-script
    # SessionState instead of the stripped closure scope.
    $applyConnectionChange = {
        param([string]$Site, [string]$Prov)
        try { Disconnect-CMSite } catch { }
        $script:ConnectedToCM = $false
        Add-LogLine -Message ('Connection settings updated: Site={0} Provider={1}' -f $Site, $Prov)
    }

    return @{
        Name    = 'Connection'
        Element = $grid
        Commit  = {
            $newSite = $txtSite.Text.Trim().ToUpper()
            $newProv = $txtProv.Text.Trim()
            # Site codes are 3 alphanumerics; anything else could corrupt the
            # WMI namespace path "root\SMS\site_$SiteCode" in downstream queries.
            # Fail closed -- throw, which the master OK handler catches and
            # surfaces via a themed error dialog.
            if ($newSite -and $newSite -notmatch '^[A-Z0-9]{3}$') {
                throw "Site Code '$newSite' is invalid. Site codes must be exactly 3 alphanumeric characters (A-Z or 0-9)."
            }
            # Provider should look like an FQDN or hostname -- whitespace, slashes,
            # quotes mean someone pasted something wrong.
            if ($newProv -and $newProv -notmatch '^[A-Za-z0-9][A-Za-z0-9\.\-]*$') {
                throw "SMS Provider '$newProv' contains invalid characters. Expected a hostname or FQDN."
            }
            $changed = ($newSite -ne [string]$global:Prefs['SiteCode']) -or ($newProv -ne [string]$global:Prefs['SMSProvider'])
            $global:Prefs['SiteCode']    = $newSite
            $global:Prefs['SMSProvider'] = $newProv
            if ($changed) {
                & $applyConnectionChange -Site $newSite -Prov $newProv
            }
        }.GetNewClosure()
    }
}

function New-LoggingPanel {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification='Panel Commit runs in a GetNewClosure scope that cannot reach script scope; prefs live in global scope for that reason (see the Preferences section).')]
    param()
    $grid = New-Object System.Windows.Controls.StackPanel
    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = 'Logging'
    $hdr.FontSize = 18
    $hdr.FontWeight = 'SemiBold'
    $hdr.Margin = '0,0,0,12'
    [void]$grid.Children.Add($hdr)

    $lblTool = New-Object System.Windows.Controls.TextBlock
    $lblTool.Text = 'Current tool log:'
    $lblTool.FontSize = 12
    $lblTool.Margin = '0,0,0,4'
    [void]$grid.Children.Add($lblTool)

    $toolTxt = New-Object System.Windows.Controls.TextBox
    $toolTxt.Name = 'txtToolLog'
    $toolTxt.Text = $toolLogPath
    $toolTxt.IsReadOnly = $true
    $toolTxt.FontSize = 11
    $toolTxt.FontFamily = 'Cascadia Code, Consolas, Courier New'
    $toolTxt.Height = 28
    $toolTxt.VerticalContentAlignment = 'Center'
    $toolTxt.Margin = '0,0,0,14'
    [void]$grid.Children.Add($toolTxt)

    $lblAudit = New-Object System.Windows.Controls.TextBlock
    $lblAudit.Text = 'Deployment audit log (JSONL):'
    $lblAudit.FontSize = 12
    $lblAudit.Margin = '0,0,0,4'
    [void]$grid.Children.Add($lblAudit)

    $row = New-Object System.Windows.Controls.DockPanel
    $row.LastChildFill = $true
    $auditBtn = New-Object System.Windows.Controls.Button
    $auditBtn.Name = 'btnBrowseAuditLog'
    $auditBtn.Content = 'Browse...'
    $auditBtn.Width = 90
    $auditBtn.Height = 28
    $auditBtn.Margin = '8,0,0,0'
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($auditBtn, [System.Windows.Controls.CharacterCasing]::Normal)
    $auditBtn.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'MahApps.Styles.Button.Square')
    [System.Windows.Controls.DockPanel]::SetDock($auditBtn, 'Right')
    [void]$row.Children.Add($auditBtn)
    $auditTxt = New-Object System.Windows.Controls.TextBox
    $auditTxt.Name = 'txtAuditLog'
    $auditTxt.Text = [string]$global:Prefs['DeploymentAuditLogPath']
    $auditTxt.FontSize = 11
    $auditTxt.FontFamily = 'Cascadia Code, Consolas, Courier New'
    $auditTxt.Height = 28
    $auditTxt.VerticalContentAlignment = 'Center'
    [void]$row.Children.Add($auditTxt)
    [void]$grid.Children.Add($row)

    $auditBtn.Add_Click({
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Title = 'Deployment audit log location'
        $dlg.Filter = 'JSONL files (*.jsonl)|*.jsonl|All files (*.*)|*.*'
        $dlg.FileName = Split-Path -Path $auditTxt.Text -Leaf
        $dir = Split-Path -Path $auditTxt.Text -Parent
        if ($dir) { $dlg.InitialDirectory = $dir }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $auditTxt.Text = $dlg.FileName
        }
    }.GetNewClosure())

    $lblRuns = New-Object System.Windows.Controls.TextBlock
    $lblRuns.Text = 'Ring run-state folder (local path or UNC share for a team):'
    $lblRuns.FontSize = 12
    $lblRuns.Margin = '0,14,0,4'
    [void]$grid.Children.Add($lblRuns)

    $runsRow = New-Object System.Windows.Controls.DockPanel
    $runsRow.LastChildFill = $true
    $runsBtn = New-Object System.Windows.Controls.Button
    $runsBtn.Name = 'btnBrowseRingRuns'
    $runsBtn.Content = 'Browse...'
    $runsBtn.Width = 90
    $runsBtn.Height = 28
    $runsBtn.Margin = '8,0,0,0'
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($runsBtn, [System.Windows.Controls.CharacterCasing]::Normal)
    $runsBtn.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'MahApps.Styles.Button.Square')
    [System.Windows.Controls.DockPanel]::SetDock($runsBtn, 'Right')
    [void]$runsRow.Children.Add($runsBtn)
    $runsTxt = New-Object System.Windows.Controls.TextBox
    $runsTxt.Name = 'txtRingRuns'
    $runsTxt.Text = [string]$global:Prefs['RingRunStatePath']
    $runsTxt.FontSize = 11
    $runsTxt.FontFamily = 'Cascadia Code, Consolas, Courier New'
    $runsTxt.Height = 28
    $runsTxt.VerticalContentAlignment = 'Center'
    [void]$runsRow.Children.Add($runsTxt)
    [void]$grid.Children.Add($runsRow)

    $runsBtn.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Folder for ring run-state files'
        if (Test-Path -LiteralPath $runsTxt.Text) { $dlg.SelectedPath = $runsTxt.Text }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $runsTxt.Text = $dlg.SelectedPath
        }
    }.GetNewClosure())

    # Commit runs after factory returns; GetNewClosure needed for $auditTxt.
    # Script-scope effects ($script:DeploymentAuditLog write, Add-LogLine call)
    # routed through a plain local scriptblock per PS51-WPF-006 so they
    # resolve against main-script SessionState, not the stripped closure.
    $applyAuditPathChange = {
        param([string]$NewPath)
        $script:DeploymentAuditLog = $NewPath
        Add-LogLine -Message ('Audit log path updated: {0}' -f $NewPath)
    }
    $applyRingRunsPathChange = {
        param([string]$NewPath)
        $script:RingRunStatePath = $NewPath
        Add-LogLine -Message ('Ring run-state folder updated: {0}' -f $NewPath)
    }

    return @{
        Name    = 'Logging'
        Element = $grid
        Commit  = {
            $newPath = $auditTxt.Text.Trim()
            if ($newPath -and $newPath -ne [string]$global:Prefs['DeploymentAuditLogPath']) {
                $global:Prefs['DeploymentAuditLogPath'] = $newPath
                & $applyAuditPathChange -NewPath $newPath
            }
            $newRuns = $runsTxt.Text.Trim()
            if ($newRuns -and $newRuns -ne [string]$global:Prefs['RingRunStatePath']) {
                $global:Prefs['RingRunStatePath'] = $newRuns
                & $applyRingRunsPathChange -NewPath $newRuns
            }
        }.GetNewClosure()
    }
}

function New-HistoryPanel {
    $grid = New-Object System.Windows.Controls.StackPanel
    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = 'Deployment History'
    $hdr.FontSize = 18
    $hdr.FontWeight = 'SemiBold'
    $hdr.Margin = '0,0,0,12'
    [void]$grid.Children.Add($hdr)

    $count = 0
    if (Test-Path -LiteralPath $script:DeploymentAuditLog) {
        try { $count = @(Get-DeploymentHistory -LogPath $script:DeploymentAuditLog).Count } catch { }
    }
    $info = New-Object System.Windows.Controls.TextBlock
    $info.Name = 'lblHistoryInfo'
    $info.Text = ('{0} record(s) in audit log.' -f $count)
    $info.FontSize = 12
    $info.Margin = '0,0,0,12'
    [void]$grid.Children.Add($info)

    $btnRow = New-Object System.Windows.Controls.StackPanel
    $btnRow.Orientation = 'Horizontal'

    $btnCsv = New-Object System.Windows.Controls.Button
    $btnCsv.Name = 'btnExportCsv'
    $btnCsv.Content = 'Export CSV...'
    $btnCsv.MinWidth = 120
    $btnCsv.Height = 28
    $btnCsv.Margin = '0,0,8,0'
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($btnCsv, [System.Windows.Controls.CharacterCasing]::Normal)
    $btnCsv.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'MahApps.Styles.Button.Square')
    [void]$btnRow.Children.Add($btnCsv)

    $btnHtml = New-Object System.Windows.Controls.Button
    $btnHtml.Name = 'btnExportHtml'
    $btnHtml.Content = 'Export HTML...'
    $btnHtml.MinWidth = 120
    $btnHtml.Height = 28
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($btnHtml, [System.Windows.Controls.CharacterCasing]::Normal)
    $btnHtml.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'MahApps.Styles.Button.Square')
    [void]$btnRow.Children.Add($btnHtml)

    [void]$grid.Children.Add($btnRow)

    # No .GetNewClosure() on these two handlers -- the body needs
    # $script:DeploymentAuditLog + $window (script scope) and calls
    # Show-ThemedMessage / Add-LogLine (script functions), and
    # GetNewClosure strips both. No factory-local captures are required
    # ($dlg / $records are declared inside the handler, module functions
    # Get-DeploymentHistory / Export-DeploymentHistoryCsv/Html resolve
    # via the module command table).
    $btnCsv.Add_Click({
        if (-not (Test-Path -LiteralPath $script:DeploymentAuditLog)) {
            [void](Show-ThemedMessage -Owner $window -Title 'Export' -Message 'No audit log exists yet.' -Buttons OK -Icon Info)
            return
        }
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Title = 'Export audit log as CSV'
        $dlg.Filter = 'CSV files (*.csv)|*.csv'
        $dlg.FileName = ('deployment-history-{0}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $records = Get-DeploymentHistory -LogPath $script:DeploymentAuditLog
            Export-DeploymentHistoryCsv -Records $records -OutputPath $dlg.FileName
            Add-LogLine -Message ('Exported CSV: {0}' -f $dlg.FileName)
        }
    })

    $btnHtml.Add_Click({
        if (-not (Test-Path -LiteralPath $script:DeploymentAuditLog)) {
            [void](Show-ThemedMessage -Owner $window -Title 'Export' -Message 'No audit log exists yet.' -Buttons OK -Icon Info)
            return
        }
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Title = 'Export audit log as HTML'
        $dlg.Filter = 'HTML files (*.html)|*.html'
        $dlg.FileName = ('deployment-history-{0}.html' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $records = Get-DeploymentHistory -LogPath $script:DeploymentAuditLog
            Export-DeploymentHistoryHtml -Records $records -OutputPath $dlg.FileName
            Add-LogLine -Message ('Exported HTML: {0}' -f $dlg.FileName)
        }
    })

    return @{
        Name    = 'History'
        Element = $grid
        Commit  = { }
    }
}

function New-TemplatesPanel {
    # Capture $window as a local var so .GetNewClosure()'d handlers inside
    # this factory can reach it. Script-scope variables ($window from
    # FindName) are stripped by GetNewClosure and evaluate to $null inside
    # the closure; local vars survive.
    $ownerWindow = $window

    $root = New-Object System.Windows.Controls.StackPanel
    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = 'Deployment Templates'
    $hdr.FontSize = 18
    $hdr.FontWeight = 'SemiBold'
    $hdr.Margin = '0,0,0,12'
    [void]$root.Children.Add($hdr)

    $templatesDir = Join-Path $PSScriptRoot 'Templates'
    if (-not (Test-Path -LiteralPath $templatesDir)) {
        New-Item -ItemType Directory -Path $templatesDir -Force | Out-Null
    }

    # Two-column: list + toolbar on the left, editor form on the right.
    $pane = New-Object System.Windows.Controls.Grid
    $pc1 = New-Object System.Windows.Controls.ColumnDefinition; $pc1.Width = '230'
    $pc2 = New-Object System.Windows.Controls.ColumnDefinition; $pc2.Width = '*'
    [void]$pane.ColumnDefinitions.Add($pc1)
    [void]$pane.ColumnDefinitions.Add($pc2)

    # -----------------------------------------------------------------------
    # LEFT: toolbar + list
    # -----------------------------------------------------------------------
    $left = New-Object System.Windows.Controls.DockPanel
    $left.LastChildFill = $true
    $left.Margin = '0,0,12,0'
    [System.Windows.Controls.Grid]::SetColumn($left, 0)

    $tb = New-Object System.Windows.Controls.StackPanel
    $tb.Orientation = 'Horizontal'
    $tb.Margin = '0,0,0,6'
    [System.Windows.Controls.DockPanel]::SetDock($tb, 'Top')

    $btnNew = New-Object System.Windows.Controls.Button
    $btnNew.Name = 'btnNew'
    $btnNew.Content = 'New'; $btnNew.MinWidth = 56; $btnNew.Height = 24; $btnNew.FontSize = 11; $btnNew.Margin = '0,0,4,0'
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($btnNew, [System.Windows.Controls.CharacterCasing]::Normal)
    $btnNew.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'MahApps.Styles.Button.Square')
    [void]$tb.Children.Add($btnNew)

    $btnDup = New-Object System.Windows.Controls.Button
    $btnDup.Name = 'btnDup'
    $btnDup.Content = 'Duplicate'; $btnDup.MinWidth = 72; $btnDup.Height = 24; $btnDup.FontSize = 11; $btnDup.Margin = '0,0,4,0'
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($btnDup, [System.Windows.Controls.CharacterCasing]::Normal)
    $btnDup.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'MahApps.Styles.Button.Square')
    [void]$tb.Children.Add($btnDup)

    $btnDel = New-Object System.Windows.Controls.Button
    $btnDel.Name = 'btnDelete'
    $btnDel.Content = 'Delete'; $btnDel.MinWidth = 56; $btnDel.Height = 24; $btnDel.FontSize = 11
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($btnDel, [System.Windows.Controls.CharacterCasing]::Normal)
    $btnDel.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'MahApps.Styles.Button.Square')
    [void]$tb.Children.Add($btnDel)

    [void]$left.Children.Add($tb)

    $list = New-Object System.Windows.Controls.ListBox
    $list.Name = 'lstTemplates'
    $list.FontSize = 12
    $list.BorderThickness = '1'
    $list.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'MahApps.Brushes.Gray8')
    [void]$left.Children.Add($list)
    [void]$pane.Children.Add($left)

    # -----------------------------------------------------------------------
    # RIGHT: editor form
    # -----------------------------------------------------------------------
    $form = New-Object System.Windows.Controls.Grid
    [System.Windows.Controls.Grid]::SetColumn($form, 1)
    $fc1 = New-Object System.Windows.Controls.ColumnDefinition; $fc1.Width = '140'
    $fc2 = New-Object System.Windows.Controls.ColumnDefinition; $fc2.Width = '*'
    $fc3 = New-Object System.Windows.Controls.ColumnDefinition; $fc3.Width = 'Auto'
    [void]$form.ColumnDefinitions.Add($fc1)
    [void]$form.ColumnDefinitions.Add($fc2)
    [void]$form.ColumnDefinitions.Add($fc3)
    0..8 | ForEach-Object {
        $rd = New-Object System.Windows.Controls.RowDefinition
        $rd.Height = 'Auto'
        [void]$form.RowDefinitions.Add($rd)
    }

    # Row 0: Name
    $lblName = New-Object System.Windows.Controls.TextBlock
    $lblName.Text = 'Name:'; $lblName.FontSize = 12; $lblName.VerticalAlignment = 'Center'; $lblName.Margin = '0,4,8,4'
    [System.Windows.Controls.Grid]::SetRow($lblName, 0); [System.Windows.Controls.Grid]::SetColumn($lblName, 0)
    [void]$form.Children.Add($lblName)
    $txtName = New-Object System.Windows.Controls.TextBox
    $txtName.Name = 'txtName'
    $txtName.FontSize = 12; $txtName.Height = 28; $txtName.VerticalContentAlignment = 'Center'; $txtName.Margin = '0,4,0,4'
    [MahApps.Metro.Controls.TextBoxHelper]::SetWatermark($txtName, 'e.g. Workstations - Pilot (Available)')
    [System.Windows.Controls.Grid]::SetRow($txtName, 0); [System.Windows.Controls.Grid]::SetColumn($txtName, 1); [System.Windows.Controls.Grid]::SetColumnSpan($txtName, 2)
    [void]$form.Children.Add($txtName)

    # Row 1: Target collection
    $lblColl = New-Object System.Windows.Controls.TextBlock
    $lblColl.Text = 'Target collection:'; $lblColl.FontSize = 12; $lblColl.VerticalAlignment = 'Center'; $lblColl.Margin = '0,4,8,4'
    [System.Windows.Controls.Grid]::SetRow($lblColl, 1); [System.Windows.Controls.Grid]::SetColumn($lblColl, 0)
    [void]$form.Children.Add($lblColl)
    $txtColl = New-Object System.Windows.Controls.TextBox
    $txtColl.Name = 'txtColl'
    $txtColl.FontSize = 12; $txtColl.Height = 28; $txtColl.VerticalContentAlignment = 'Center'; $txtColl.Margin = '0,4,0,4'
    [MahApps.Metro.Controls.TextBoxHelper]::SetWatermark($txtColl, 'Device collection name')
    [System.Windows.Controls.Grid]::SetRow($txtColl, 1); [System.Windows.Controls.Grid]::SetColumn($txtColl, 1)
    [void]$form.Children.Add($txtColl)
    $btnBrowseColl = New-Object System.Windows.Controls.Button
    $btnBrowseColl.Name = 'btnBrowseColl'
    $btnBrowseColl.Content = 'Browse...'; $btnBrowseColl.Width = 90; $btnBrowseColl.Height = 28; $btnBrowseColl.Margin = '8,4,0,4'
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($btnBrowseColl, [System.Windows.Controls.CharacterCasing]::Normal)
    $btnBrowseColl.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'MahApps.Styles.Button.Square')
    [System.Windows.Controls.Grid]::SetRow($btnBrowseColl, 1); [System.Windows.Controls.Grid]::SetColumn($btnBrowseColl, 2)
    [void]$form.Children.Add($btnBrowseColl)

    # Row 2: Purpose
    $lblPurp = New-Object System.Windows.Controls.TextBlock
    $lblPurp.Text = 'Purpose:'; $lblPurp.FontSize = 12; $lblPurp.VerticalAlignment = 'Center'; $lblPurp.Margin = '0,4,8,4'
    [System.Windows.Controls.Grid]::SetRow($lblPurp, 2); [System.Windows.Controls.Grid]::SetColumn($lblPurp, 0)
    [void]$form.Children.Add($lblPurp)
    $pnlPurp = New-Object System.Windows.Controls.StackPanel
    $pnlPurp.Orientation = 'Horizontal'; $pnlPurp.Margin = '0,4,0,4'
    $rAvail = New-Object System.Windows.Controls.RadioButton
    $rAvail.Name = 'rAvail'
    $rAvail.Content = 'Available'; $rAvail.GroupName = 'TplPurpose'; $rAvail.IsChecked = $true
    $rAvail.FontSize = 12; $rAvail.VerticalAlignment = 'Center'; $rAvail.Margin = '0,0,16,0'
    [void]$pnlPurp.Children.Add($rAvail)
    $rReq = New-Object System.Windows.Controls.RadioButton
    $rReq.Name = 'rReq'
    $rReq.Content = 'Required'; $rReq.GroupName = 'TplPurpose'
    $rReq.FontSize = 12; $rReq.VerticalAlignment = 'Center'
    [void]$pnlPurp.Children.Add($rReq)
    [System.Windows.Controls.Grid]::SetRow($pnlPurp, 2); [System.Windows.Controls.Grid]::SetColumn($pnlPurp, 1); [System.Windows.Controls.Grid]::SetColumnSpan($pnlPurp, 2)
    [void]$form.Children.Add($pnlPurp)

    # Row 3: Notification
    $lblNot = New-Object System.Windows.Controls.TextBlock
    $lblNot.Text = 'Notification:'; $lblNot.FontSize = 12; $lblNot.VerticalAlignment = 'Center'; $lblNot.Margin = '0,4,8,4'
    [System.Windows.Controls.Grid]::SetRow($lblNot, 3); [System.Windows.Controls.Grid]::SetColumn($lblNot, 0)
    [void]$form.Children.Add($lblNot)
    $cboNot = New-Object System.Windows.Controls.ComboBox
    $cboNot.Name = 'cboNot'
    $cboNot.FontSize = 12; $cboNot.Height = 28; $cboNot.VerticalContentAlignment = 'Center'; $cboNot.Margin = '0,4,0,4'
    foreach ($n in @(
        @{C='Display All'; T='DisplayAll'},
        @{C='Display in Software Center Only'; T='DisplaySoftwareCenterOnly'},
        @{C='Hide All'; T='HideAll'}
    )) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $n.C; $item.Tag = $n.T
        [void]$cboNot.Items.Add($item)
    }
    $cboNot.SelectedIndex = 0
    [System.Windows.Controls.Grid]::SetRow($cboNot, 3); [System.Windows.Controls.Grid]::SetColumn($cboNot, 1); [System.Windows.Controls.Grid]::SetColumnSpan($cboNot, 2)
    [void]$form.Children.Add($cboNot)

    # Row 4: Time basis
    $lblTime = New-Object System.Windows.Controls.TextBlock
    $lblTime.Text = 'Time basis:'; $lblTime.FontSize = 12; $lblTime.VerticalAlignment = 'Center'; $lblTime.Margin = '0,4,8,4'
    [System.Windows.Controls.Grid]::SetRow($lblTime, 4); [System.Windows.Controls.Grid]::SetColumn($lblTime, 0)
    [void]$form.Children.Add($lblTime)
    $pnlTime = New-Object System.Windows.Controls.StackPanel
    $pnlTime.Orientation = 'Horizontal'; $pnlTime.Margin = '0,4,0,4'
    $rLocal = New-Object System.Windows.Controls.RadioButton
    $rLocal.Name = 'rLocal'
    $rLocal.Content = 'Local time'; $rLocal.GroupName = 'TplTime'; $rLocal.IsChecked = $true
    $rLocal.FontSize = 12; $rLocal.VerticalAlignment = 'Center'; $rLocal.Margin = '0,0,16,0'
    [void]$pnlTime.Children.Add($rLocal)
    $rUtc = New-Object System.Windows.Controls.RadioButton
    $rUtc.Name = 'rUtc'
    $rUtc.Content = 'UTC'; $rUtc.GroupName = 'TplTime'
    $rUtc.FontSize = 12; $rUtc.VerticalAlignment = 'Center'
    [void]$pnlTime.Children.Add($rUtc)
    [System.Windows.Controls.Grid]::SetRow($pnlTime, 4); [System.Windows.Controls.Grid]::SetColumn($pnlTime, 1); [System.Windows.Controls.Grid]::SetColumnSpan($pnlTime, 2)
    [void]$form.Children.Add($pnlTime)

    # Row 5: Deadline offset hours
    $lblDead = New-Object System.Windows.Controls.TextBlock
    $lblDead.Text = 'Deadline offset (h):'; $lblDead.FontSize = 12; $lblDead.VerticalAlignment = 'Center'; $lblDead.Margin = '0,4,8,4'
    [System.Windows.Controls.Grid]::SetRow($lblDead, 5); [System.Windows.Controls.Grid]::SetColumn($lblDead, 0)
    [void]$form.Children.Add($lblDead)
    $txtDead = New-Object System.Windows.Controls.TextBox
    $txtDead.Name = 'txtDead'
    $txtDead.FontSize = 12; $txtDead.Height = 28; $txtDead.VerticalContentAlignment = 'Center'; $txtDead.Margin = '0,4,0,4'
    $txtDead.Width = 80; $txtDead.HorizontalAlignment = 'Left'; $txtDead.Text = '0'
    [System.Windows.Controls.Grid]::SetRow($txtDead, 5); [System.Windows.Controls.Grid]::SetColumn($txtDead, 1)
    [void]$form.Children.Add($txtDead)

    # Row 6: Checkboxes (WrapPanel so they flow on narrow editors)
    $chkPane = New-Object System.Windows.Controls.WrapPanel
    $chkPane.Margin = '0,6,0,0'
    $chkOverride = New-Object System.Windows.Controls.CheckBox; $chkOverride.Name = 'chkOverride'; $chkOverride.Content = 'Override Service Window'; $chkOverride.FontSize = 12; $chkOverride.Margin = '0,0,12,6'; [void]$chkPane.Children.Add($chkOverride)
    $chkReboot   = New-Object System.Windows.Controls.CheckBox; $chkReboot.Name = 'chkReboot';     $chkReboot.Content   = 'Reboot outside SW';       $chkReboot.FontSize   = 12; $chkReboot.Margin   = '0,0,12,6'; [void]$chkPane.Children.Add($chkReboot)
    $chkMetered  = New-Object System.Windows.Controls.CheckBox; $chkMetered.Name = 'chkMetered';   $chkMetered.Content  = 'Allow metered';           $chkMetered.FontSize  = 12; $chkMetered.Margin  = '0,0,12,6'; [void]$chkPane.Children.Add($chkMetered)
    $chkBoundary = New-Object System.Windows.Controls.CheckBox; $chkBoundary.Name = 'chkBoundary'; $chkBoundary.Content = 'Boundary fallback';       $chkBoundary.FontSize = 12; $chkBoundary.Margin = '0,0,12,6'; $chkBoundary.IsChecked = $true; [void]$chkPane.Children.Add($chkBoundary)
    $chkMsUpd    = New-Object System.Windows.Controls.CheckBox; $chkMsUpd.Name = 'chkMsUpd';       $chkMsUpd.Content    = 'Allow MS Update';         $chkMsUpd.FontSize    = 12; $chkMsUpd.Margin    = '0,0,12,6'; [void]$chkPane.Children.Add($chkMsUpd)
    $chkFull     = New-Object System.Windows.Controls.CheckBox; $chkFull.Name = 'chkFull';         $chkFull.Content     = 'Full scan post-reboot';   $chkFull.FontSize     = 12; $chkFull.Margin     = '0,0,12,6'; $chkFull.IsChecked = $true; [void]$chkPane.Children.Add($chkFull)
    [System.Windows.Controls.Grid]::SetRow($chkPane, 6); [System.Windows.Controls.Grid]::SetColumn($chkPane, 0); [System.Windows.Controls.Grid]::SetColumnSpan($chkPane, 3)
    [void]$form.Children.Add($chkPane)

    # Row 7: Save button
    $btnSave = New-Object System.Windows.Controls.Button
    $btnSave.Name = 'btnSave'
    $btnSave.Content = 'Save template'; $btnSave.MinWidth = 130; $btnSave.Height = 28
    $btnSave.HorizontalAlignment = 'Left'; $btnSave.Margin = '0,12,0,0'
    [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($btnSave, [System.Windows.Controls.CharacterCasing]::Normal)
    $btnSave.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, 'MahApps.Styles.Button.Square.Accent')
    [System.Windows.Controls.Grid]::SetRow($btnSave, 7); [System.Windows.Controls.Grid]::SetColumn($btnSave, 0); [System.Windows.Controls.Grid]::SetColumnSpan($btnSave, 3)
    [void]$form.Children.Add($btnSave)

    # Row 8: Status
    $lblStatus = New-Object System.Windows.Controls.TextBlock
    $lblStatus.Name = 'lblTemplatesStatus'
    $lblStatus.FontSize = 11; $lblStatus.Margin = '0,8,0,0'; $lblStatus.TextWrapping = 'Wrap'
    $lblStatus.Foreground = if ($toggleTheme.IsOn) { $script:LogLabelDark } else { $script:LogLabelLight }
    [System.Windows.Controls.Grid]::SetRow($lblStatus, 8); [System.Windows.Controls.Grid]::SetColumn($lblStatus, 0); [System.Windows.Controls.Grid]::SetColumnSpan($lblStatus, 3)
    [void]$form.Children.Add($lblStatus)

    [void]$pane.Children.Add($form)
    [void]$root.Children.Add($pane)

    # -----------------------------------------------------------------------
    # Local state (per-panel instance) -- captured by handler closures
    # -----------------------------------------------------------------------
    # Holds @{Name; FilePath; Data} entries in the same order as $list.Items.
    $state = @{ Items = New-Object System.Collections.ArrayList }

    # -----------------------------------------------------------------------
    # Helper scriptblocks (LOCAL variables -- NOT $script:).
    #
    # Two scope-strip rules at play:
    #   1. .GetNewClosure() on a WPF handler (invoked via InvokeAsDelegateHelper)
    #      strips script-scope function lookup AND $script: variable access.
    #   2. A scriptblock defined as a LOCAL variable in this factory carries
    #      the factory's SessionState (script scope). Invoking it via `&`
    #      from inside a GetNewClosure'd handler runs the body in its OWN
    #      SessionState -- so script functions and $script: vars resolve
    #      correctly from within the helper body.
    #
    # Therefore: anything a handler needs to call that lives outside the
    # local factory scope (script functions like Show-ThemedMessage, or
    # $script: vars like $script:RefreshApplyTemplateCombo) gets a local
    # helper defined here, and the handler calls `& $helper`.
    #
    # The helpers below (sanitize/refreshList/loadEditor/harvestConfig) use
    # .GetNewClosure() because they reference the local control refs
    # ($txtName, $list, etc.) created in this factory. The "call out to
    # script scope" helpers (showThemed, showCollectionBrowse, connectIfNeeded,
    # refreshMainCombo) are PLAIN (no GetNewClosure) so they keep their
    # factory SessionState and can resolve script-scope callees.
    # -----------------------------------------------------------------------
    $showThemed = {
        param(
            [Parameter(Mandatory)]$Owner,
            [Parameter(Mandatory)][string]$Title,
            [Parameter(Mandatory)][string]$Message,
            [ValidateSet('OK','OKCancel','YesNo')][string]$Buttons = 'OK',
            [ValidateSet('None','Info','Warn','Error','Question')][string]$Icon = 'None'
        )
        Show-ThemedMessage -Owner $Owner -Title $Title -Message $Message -Buttons $Buttons -Icon $Icon
    }

    $showCollectionBrowse = {
        param([Parameter(Mandatory)]$Owner, [switch]$Force)
        Show-CollectionBrowse -Owner $Owner -Force:$Force
    }

    $connectIfNeeded  = { Connect-IfNeeded }
    $refreshMainCombo = {
        if ($null -ne $script:RefreshApplyTemplateCombo) { & $script:RefreshApplyTemplateCombo }
    }

    $sanitize = {
        param($name)
        $s = ($name -replace '[\\/:*?"<>|]', '_').Trim()
        if ([string]::IsNullOrWhiteSpace($s)) { $s = 'Untitled' }
        return $s
    }.GetNewClosure()

    $refreshList = {
        $list.Items.Clear()
        $state.Items.Clear() | Out-Null
        if (Test-Path -LiteralPath $templatesDir) {
            $files = @(Get-ChildItem -LiteralPath $templatesDir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object Name)
            foreach ($f in $files) {
                try {
                    $data = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json
                    $name = if ($data.Name) { [string]$data.Name } else { [System.IO.Path]::GetFileNameWithoutExtension($f.Name) }
                    $state.Items.Add(@{ Name = $name; FilePath = $f.FullName; Data = $data }) | Out-Null
                    [void]$list.Items.Add($name)
                } catch {
                    Write-Log ('Skipping unreadable template ' + $f.Name + ': ' + $_) -Level WARN
                }
            }
        }
        $lblStatus.Text = ('{0} template(s) in ' + $templatesDir) -f $state.Items.Count
    }.GetNewClosure()

    $loadEditor = {
        param($data)
        $txtName.Text  = [string]$data.Name
        $txtColl.Text  = [string]$data.TargetCollectionName
        if (([string]$data.DeployPurpose) -eq 'Required') { $rReq.IsChecked = $true } else { $rAvail.IsChecked = $true }
        $target = [string]$data.UserNotification
        $sel = 0
        for ($i = 0; $i -lt $cboNot.Items.Count; $i++) {
            if ([string]$cboNot.Items[$i].Tag -eq $target) { $sel = $i; break }
        }
        $cboNot.SelectedIndex = $sel
        if (([string]$data.TimeBasedOn) -eq 'UTC') { $rUtc.IsChecked = $true } else { $rLocal.IsChecked = $true }
        $txtDead.Text = [string]$(if ($null -ne $data.DefaultDeadlineOffsetHours) { $data.DefaultDeadlineOffsetHours } else { 0 })
        $chkOverride.IsChecked = [bool]$data.OverrideServiceWindow
        $chkReboot.IsChecked   = [bool]$data.RebootOutsideServiceWindow
        $chkMetered.IsChecked  = [bool]$data.AllowMeteredConnection
        $chkBoundary.IsChecked = if ($null -ne $data.AllowBoundaryFallback) { [bool]$data.AllowBoundaryFallback } else { $true }
        $chkMsUpd.IsChecked    = [bool]$data.AllowMicrosoftUpdate
        $chkFull.IsChecked     = if ($null -ne $data.RequirePostRebootFullScan) { [bool]$data.RequirePostRebootFullScan } else { $true }
    }.GetNewClosure()

    $harvestConfig = {
        $purpose   = if ($rReq.IsChecked) { 'Required' } else { 'Available' }
        $timeBasis = if ($rUtc.IsChecked) { 'UTC' } else { 'LocalTime' }
        $notif     = 'DisplayAll'
        if ($cboNot.SelectedItem -and $cboNot.SelectedItem.Tag) { $notif = [string]$cboNot.SelectedItem.Tag }
        $hours = 0
        [int]::TryParse($txtDead.Text.Trim(), [ref]$hours) | Out-Null
        return @{
            TargetCollectionName       = $txtColl.Text.Trim()
            TargetCollectionID         = ''
            DeployPurpose              = $purpose
            UserNotification           = $notif
            TimeBasedOn                = $timeBasis
            OverrideServiceWindow      = [bool]$chkOverride.IsChecked
            RebootOutsideServiceWindow = [bool]$chkReboot.IsChecked
            AllowMeteredConnection     = [bool]$chkMetered.IsChecked
            AllowBoundaryFallback      = [bool]$chkBoundary.IsChecked
            AllowMicrosoftUpdate       = [bool]$chkMsUpd.IsChecked
            RequirePostRebootFullScan  = [bool]$chkFull.IsChecked
            DefaultDeadlineOffsetHours = $hours
        }
    }.GetNewClosure()

    # -----------------------------------------------------------------------
    # Handlers
    # -----------------------------------------------------------------------
    $list.Add_SelectionChanged({
        if ($list.SelectedIndex -lt 0) { return }
        $t = $state.Items[$list.SelectedIndex]
        if ($t) { & $loadEditor $t.Data }
    }.GetNewClosure())

    $btnBrowseColl.Add_Click({
        if (-not (& $connectIfNeeded)) { return }
        $reload = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift) -ne 0
        $picked = & $showCollectionBrowse -Owner $ownerWindow -Force:$reload
        if ($picked) { $txtColl.Text = $picked }
    }.GetNewClosure())

    $btnNew.Add_Click({
        $defaults = [pscustomobject]@{
            Name = 'New Template'
            TargetCollectionName = ''
            TargetCollectionID = ''
            DeployPurpose = 'Available'
            UserNotification = 'DisplayAll'
            TimeBasedOn = 'LocalTime'
            OverrideServiceWindow = $false
            RebootOutsideServiceWindow = $false
            AllowMeteredConnection = $false
            AllowBoundaryFallback = $true
            AllowMicrosoftUpdate = $false
            RequirePostRebootFullScan = $true
            DefaultDeadlineOffsetHours = 0
        }
        $list.SelectedIndex = -1
        & $loadEditor $defaults
        $txtName.Focus() | Out-Null
        $txtName.SelectAll()
        $lblStatus.Text = 'New template -- edit fields then click Save.'
    }.GetNewClosure())

    $btnDup.Add_Click({
        if ($list.SelectedIndex -lt 0) { return }
        $t = $state.Items[$list.SelectedIndex]
        if (-not $t) { return }
        # Build a copy of the data object with a renamed Name
        $copy = [pscustomobject]@{}
        foreach ($p in $t.Data.PSObject.Properties) {
            Add-Member -InputObject $copy -NotePropertyName $p.Name -NotePropertyValue $p.Value
        }
        $copy.Name = ($t.Name + ' (copy)')
        $list.SelectedIndex = -1
        & $loadEditor $copy
        $txtName.Focus() | Out-Null
        $lblStatus.Text = 'Duplicate -- rename and click Save.'
    }.GetNewClosure())

    $btnDel.Add_Click({
        if ($list.SelectedIndex -lt 0) { return }
        $t = $state.Items[$list.SelectedIndex]
        if (-not $t) { return }
        $ans = & $showThemed -Owner $ownerWindow -Title 'Delete template' `
            -Message ("Delete template '{0}'?" -f $t.Name) -Buttons YesNo -Icon Question
        if ($ans -ne 'Yes') { return }
        try {
            Remove-DeploymentTemplate -TemplatePath $t.FilePath
            Write-Log ('Template deleted: ' + $t.Name)
            & $refreshList
            & $refreshMainCombo
        } catch {
            [void](& $showThemed -Owner $ownerWindow -Title 'Delete failed' -Message ($_.ToString()) -Buttons OK -Icon Error)
        }
    }.GetNewClosure())

    $btnSave.Add_Click({
        $name = $txtName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($name)) {
            [void](& $showThemed -Owner $ownerWindow -Title 'Save template' -Message 'Template name cannot be empty.' -Buttons OK -Icon Warn)
            return
        }
        $config   = & $harvestConfig
        $fileBase = & $sanitize $name
        $filePath = Join-Path $templatesDir ($fileBase + '.json')

        # On rename: sanitized name produces a new filename, so the old file stays
        # behind if we don't prune it. Only prune when the user had a previous
        # file loaded and the target path differs.
        if ($list.SelectedIndex -ge 0) {
            $cur = $state.Items[$list.SelectedIndex]
            if ($cur -and $cur.FilePath -and ($cur.FilePath -ne $filePath) -and (Test-Path -LiteralPath $cur.FilePath)) {
                Remove-Item -LiteralPath $cur.FilePath -Force
            }
        }

        try {
            Save-DeploymentTemplate -TemplatePath $filePath -TemplateName $name -Config $config
            Write-Log ('Template saved: ' + $name)
            & $refreshList
            for ($i = 0; $i -lt $list.Items.Count; $i++) {
                if ([string]$list.Items[$i] -eq $name) { $list.SelectedIndex = $i; break }
            }
            & $refreshMainCombo
        } catch {
            [void](& $showThemed -Owner $ownerWindow -Title 'Save failed' -Message ($_.ToString()) -Buttons OK -Icon Error)
        }
    }.GetNewClosure())

    # Initial populate
    & $refreshList

    return @{
        Name    = 'Templates'
        Element = $root
        Commit  = { }
    }
}

function New-AboutPanel {
    $grid = New-Object System.Windows.Controls.StackPanel
    $hdr = New-Object System.Windows.Controls.TextBlock
    $hdr.Text = 'About'
    $hdr.FontSize = 18
    $hdr.FontWeight = 'SemiBold'
    $hdr.Margin = '0,0,0,12'
    [void]$grid.Children.Add($hdr)

    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = 'Deployment Helper'
    $title.FontSize = 20
    $title.FontWeight = 'Bold'
    $title.Margin = '0,0,0,4'
    [void]$grid.Children.Add($title)

    $ver = New-Object System.Windows.Controls.TextBlock
    $ver.Text = 'Version ' + $script:AppVersion
    $ver.FontSize = 12
    $ver.Margin = '0,0,0,12'
    [void]$grid.Children.Add($ver)

    $desc = New-Object System.Windows.Controls.TextBlock
    $desc.Text = 'Safe Configuration Manager deployment for Apps, Packages, Task Sequences, and Software Update Groups with 5-check validation, ring deployments, and audit logging.'
    $desc.FontSize = 12
    $desc.TextWrapping = 'Wrap'
    $desc.Margin = '0,0,0,12'
    [void]$grid.Children.Add($desc)

    foreach ($line in @('Author: jasonulbright', 'License: MIT', 'Repo: https://github.com/jasonulbright/deployment-helper')) {
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text = $line
        $tb.FontSize = 11
        $tb.Margin = '0,0,0,4'
        [void]$grid.Children.Add($tb)
    }

    return @{
        Name    = 'About'
        Element = $grid
        Commit  = { }
    }
}

function Show-OptionsDialog {
    param(
        [Parameter(Mandatory)]$Owner,
        [string]$InitialSection = 'Connection'
    )

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Options"
    Width="860" Height="580"
    MinWidth="760" MinHeight="480"
    WindowStartupLocation="CenterOwner"
    TitleCharacterCasing="Normal"
    ShowIconOnTitleBar="False"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="210"/>
            <ColumnDefinition Width="1"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <ListBox Grid.Column="0" Grid.Row="0" x:Name="lstNav" BorderThickness="0" Padding="0,8,0,0">
            <ListBox.ItemContainerStyle>
                <Style TargetType="ListBoxItem">
                    <Setter Property="Padding" Value="16,10,16,10"/>
                    <Setter Property="FontSize" Value="13"/>
                </Style>
            </ListBox.ItemContainerStyle>
        </ListBox>

        <Border Grid.Column="1" Grid.Row="0" Background="{DynamicResource MahApps.Brushes.Gray8}"/>

        <ScrollViewer Grid.Column="2" Grid.Row="0" VerticalScrollBarVisibility="Auto">
            <ContentControl x:Name="contentArea" Margin="20,18,20,18"/>
        </ScrollViewer>

        <Border Grid.Column="0" Grid.ColumnSpan="3" Grid.Row="1"
                BorderBrush="{DynamicResource MahApps.Brushes.Gray8}" BorderThickness="0,1,0,0">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="20,12,20,12">
                <Button x:Name="btnOK"     Content="OK"     MinWidth="90" Height="32" Margin="0,0,8,0" IsDefault="True"
                        Style="{DynamicResource MahApps.Styles.Button.Square.Accent}"
                        Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
                <Button x:Name="btnCancel" Content="Cancel" MinWidth="90" Height="32" IsCancel="True"
                        Style="{DynamicResource MahApps.Styles.Button.Square}"
                        Controls:ControlsHelper.ContentCharacterCasing="Normal"/>
            </StackPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@

    [xml]$xml = $dlgXaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    $dlg    = [System.Windows.Markup.XamlReader]::Load($reader)
    Install-TitleBarDragFallback -Window $dlg

    $theme = [ControlzEx.Theming.ThemeManager]::Current.DetectTheme($Owner)
    if ($theme) { [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($dlg, $theme) }
    $dlg.Owner = $Owner
    try {
        $dlg.WindowTitleBrush          = $Owner.WindowTitleBrush
        $dlg.NonActiveWindowTitleBrush = $Owner.WindowTitleBrush
        $dlg.GlowBrush                 = $Owner.GlowBrush
        $dlg.NonActiveGlowBrush        = $Owner.GlowBrush
    } catch { }

    $lstNav      = $dlg.FindName('lstNav')
    $contentArea = $dlg.FindName('contentArea')
    $btnOK       = $dlg.FindName('btnOK')
    $btnCancel   = $dlg.FindName('btnCancel')

    $panels = @(
        (New-ConnectionPanel),
        (New-LoggingPanel),
        (New-HistoryPanel),
        (New-TemplatesPanel),
        (New-AboutPanel)
    )

    foreach ($p in $panels) { [void]$lstNav.Items.Add($p.Name) }

    # No .GetNewClosure() on these handlers: Show-OptionsDialog is still
    # blocked on $dlg.ShowDialog() when they fire, so $panels / $contentArea /
    # $dlg / $btnOK reach lexical parent scope naturally -- AND so do the
    # script-level helpers (Save-DhPreferences, Show-ThemedMessage) that
    # these handlers call. GetNewClosure would strip script-function lookup
    # via ScriptBlock.InvokeAsDelegateHelper, same pattern as the Session 7
    # main-window Loaded fix.
    $lstNav.Add_SelectionChanged({
        $idx = $lstNav.SelectedIndex
        if ($idx -ge 0 -and $idx -lt $panels.Count) {
            $contentArea.Content = $panels[$idx].Element
        }
    })

    $initialIdx = 0
    for ($i = 0; $i -lt $panels.Count; $i++) {
        if ($panels[$i].Name -eq $InitialSection) { $initialIdx = $i; break }
    }
    $lstNav.SelectedIndex = $initialIdx

    $script:OptionsDlgResult = $false
    $btnOK.Add_Click({
        try {
            foreach ($p in $panels) { if ($p.Commit) { & $p.Commit } }
            Save-DhPreferences -Prefs $global:Prefs
            $script:OptionsDlgResult = $true
            $dlg.Close()
        } catch {
            [void](Show-ThemedMessage -Owner $dlg -Title 'Save failed' -Message ('Save failed: {0}' -f $_.Exception.Message) -Buttons OK -Icon Error)
        }
    })

    $btnCancel.Add_Click({ $dlg.Close() })

    [void]$dlg.ShowDialog()

    if ($script:OptionsDlgResult) {
        Add-LogLine -Message 'Options saved.'
    }
}

# =============================================================================
# Deployment type switcher
# =============================================================================
$script:SetCurrentType = {
    param([string]$Type)

    $script:CurrentType = $Type
    $formPane.Visibility         = [System.Windows.Visibility]::Visible
    $ringPane.Visibility         = [System.Windows.Visibility]::Collapsed
    $pnlApplyTemplate.Visibility = [System.Windows.Visibility]::Visible
    $meta = $script:TypeMeta[$Type]
    $txtModuleHeader.Text    = $meta.Header
    $txtModuleSubheader.Text = $meta.Subheader
    $lblTargetName.Text      = $meta.TargetLabel
    $lblCheck1.Text          = $meta.Check1Text
    [MahApps.Metro.Controls.TextBoxHelper]::SetWatermark($txtTargetName, $meta.Watermark)

    # Program row: Packages only
    $showProgram = ($Type -eq 'Packages')
    $progVis = if ($showProgram) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $lblProgram.Visibility = $progVis
    $cboProgram.Visibility = $progVis
    if (-not $showProgram) { $cboProgram.Items.Clear() }

    # Notification row: Apps + TS + SUG use it with different item sets / labels; Packages hides it
    switch ($Type) {
        'Apps' {
            Set-NotificationItems -Items $script:AppsNotificationItems -LabelText 'Notification:'
            $lblNotification.Visibility = [System.Windows.Visibility]::Visible
            $cboNotification.Visibility = [System.Windows.Visibility]::Visible
        }
        'TaskSequences' {
            Set-NotificationItems -Items $script:TaskSequenceAvailabilityItems -LabelText 'Availability:'
            $lblNotification.Visibility = [System.Windows.Visibility]::Visible
            $cboNotification.Visibility = [System.Windows.Visibility]::Visible
        }
        'SUG' {
            Set-NotificationItems -Items $script:AppsNotificationItems -LabelText 'Notification:'
            $lblNotification.Visibility = [System.Windows.Visibility]::Visible
            $cboNotification.Visibility = [System.Windows.Visibility]::Visible
        }
        default {
            $lblNotification.Visibility = [System.Windows.Visibility]::Collapsed
            $cboNotification.Visibility = [System.Windows.Visibility]::Collapsed
        }
    }

    # Grid.Row=9 is shared by three mutually-exclusive panels:
    #   pnlRequiredOptions      -- Apps + Packages, Purpose=Required only
    #   pnlTaskSequenceOptions  -- TaskSequences (Purpose-independent)
    #   pnlSUGOptions           -- SUG (Purpose-independent)
    # Earlier comment claimed "gating is orthogonal" but that's wrong:
    # Type and Purpose can coexist (e.g. TS + Required), and without
    # explicit gating both the TS panel and the Required panel would
    # render stacked in the same row. Resolve here: pick exactly one
    # panel based on (Type, Purpose) and collapse the other two.
    $showTs       = ($Type -eq 'TaskSequences')
    $showSug      = ($Type -eq 'SUG')
    $showPkg      = ($Type -eq 'Packages')
    # pnlRequiredOptions is now Apps-only. Packages has its own
    # pnlPackageOptions that includes the Required-extras checkboxes
    # scoped to chkPkg* so there's no Grid.Row=9 overlap.
    $showRequired = ($Type -eq 'Apps') -and ($radRequired.IsChecked -eq $true)
    $pnlTaskSequenceOptions.Visibility = if ($showTs)       { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $pnlSUGOptions.Visibility          = if ($showSug)      { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $pnlPackageOptions.Visibility      = if ($showPkg)      { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $pnlRequiredOptions.Visibility     = if ($showRequired) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }

    # Distribution row: Apps + Packages + TaskSequences. SUG excluded -- update content
    # for SUGs lives in a separate deployment package flow (v1.1 candidate).
    $distVis = if ($Type -in @('Apps','Packages','TaskSequences')) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $lblDistribute.Visibility = $distVis
    $pnlDistribute.Visibility = $distVis
    # Reset selection when type changes -- different target, different DP picks
    $global:SelectedDPGroups = @()
    Update-DPPickerButtonLabel

    # Validate button: all four types wired through session 5
    $btnValidate.Visibility = [System.Windows.Visibility]::Visible

    Reset-ValidationUi
    Set-StatusText -Text ('Viewing {0}.' -f $meta.Header)
}

# =============================================================================
# Ring deployment view
#
# One object, one ring plan, N ordinary deployments. The preview grid holds
# the expanded rings for this run only; plan files are never written here.
# Create-all makes every ring now with future dates. Hold later rings writes
# a run file, creates ring 1, and leaves the rest to Promote in Open runs.
# =============================================================================
$script:RingPlans           = @()
$script:RingPlanErrors      = @()
$script:RingPreviewRows     = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
$script:RingValidated       = $null
$script:RingSuppress        = $false
$script:RingViewInitialized = $false
$dgRingPreview.ItemsSource  = $script:RingPreviewRows

$script:RingTypeBrowse = @{
    'Application'  = @{ Browse = 'Apps';          NameProperty = 'LocalizedDisplayName'; Title = 'Browse applications';           Watermark = 'Filter by name, version, or CI ID' }
    'Package'      = @{ Browse = 'Packages';      NameProperty = 'Name';                 Title = 'Browse packages';               Watermark = 'Filter by name, package ID, manufacturer, or version' }
    'TaskSequence' = @{ Browse = 'TaskSequences'; NameProperty = 'Name';                 Title = 'Browse task sequences';         Watermark = 'Filter by name, package ID, or description' }
    'SUG'          = @{ Browse = 'SUG';           NameProperty = 'LocalizedDisplayName'; Title = 'Browse software update groups'; Watermark = 'Filter by name' }
}

function Get-RingSelectedType {
    if ($cboRingType.SelectedItem) { return [string]$cboRingType.SelectedItem.Tag }
    return 'Application'
}

function Get-RingSelectedPlan {
    $i = $cboRingPlan.SelectedIndex
    $plans = @($script:RingPlans)
    if ($i -ge 0 -and $i -lt $plans.Count) { return $plans[$i] }
    return $null
}

function Format-RingDate {
    param($Value)
    try { $d = ConvertFrom-RingDateText $Value } catch { return [string]$Value }
    if ($null -eq $d) { return '' }
    return $d.ToString('yyyy-MM-dd HH:mm')
}

function Reset-RingValidation {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates window controls only.')]
    param()
    $script:RingValidated = $null
    $btnRingCreate.Visibility = [System.Windows.Visibility]::Collapsed
}

function Update-RingPreviewGrid {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates window controls only.')]
    param()
    # Items.Refresh throws while a cell is still in edit mode.
    try { [void]$dgRingPreview.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) } catch { $null = $_ }
    $dgRingPreview.Items.Refresh()
}

function Update-RingPreview {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates window controls only.')]
    param()
    Reset-RingValidation
    $script:RingPreviewRows.Clear()
    $plan = Get-RingSelectedPlan
    if (-not $plan) { return }
    $start = if ($dtpRingStart.SelectedDateTime) { [datetime]$dtpRingStart.SelectedDateTime } else { Get-Date }
    foreach ($r in (Expand-RingPlan -Plan $plan -Start $start)) {
        $r | Add-Member -NotePropertyName AvailableText -NotePropertyValue (Format-RingDate $r.AvailableDateTime)
        $r | Add-Member -NotePropertyName DeadlineText  -NotePropertyValue (Format-RingDate $r.DeadlineDateTime)
        $r | Add-Member -NotePropertyName ThresholdText -NotePropertyValue $(if ($null -ne $r.SuccessThresholdPercent) { ('{0}%' -f $r.SuccessThresholdPercent) } else { '' })
        $r | Add-Member -NotePropertyName Checks        -NotePropertyValue ''
        $r | Add-Member -NotePropertyName MemberCount   -NotePropertyValue $null
        $script:RingPreviewRows.Add($r)
    }
    Update-RingPreviewGrid
}

function Update-RingPlanSelection {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates window controls only.')]
    param()
    $plan = Get-RingSelectedPlan
    $refused = @($script:RingPlanErrors).Count
    if (-not $plan) {
        $txtRingPlanInfo.Text = if ($refused -gt 0) {
            ('No plan loaded. {0} plan file(s) in {1} were refused; see the log.' -f $refused, $script:RingPlanDir)
        } else {
            ('No plan files in {0}.' -f $script:RingPlanDir)
        }
        Reset-RingValidation
        $script:RingPreviewRows.Clear()
        return
    }
    $info = ('{0} File: {1}. Time basis: {2}.' -f $plan.Description, $plan.FilePath, $plan.TimeBasedOn).Trim()
    if ($plan.TimeBasedOn -eq 'Utc') { $info += ' Enter the plan start and the preview times in UTC.' }
    if ($refused -gt 0) { $info += (' {0} other plan file(s) were refused; see the log.' -f $refused) }
    $txtRingPlanInfo.Text = $info
    $chkRingHold.IsChecked = [bool]$plan.HoldLaterRings
    Update-RingPreview
}

function Update-RingPlanCombo {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates window controls only.')]
    param()
    $current = Get-RingSelectedPlan
    $previous = if ($current) { $current.Name } else { $null }
    $script:RingSuppress = $true
    try {
        $list = Get-RingPlanList -Path $script:RingPlanDir
        $script:RingPlans      = @($list.Plans)
        $script:RingPlanErrors = @($list.Errors)
        $cboRingPlan.Items.Clear()
        $selected = 0
        for ($i = 0; $i -lt $script:RingPlans.Count; $i++) {
            $p = $script:RingPlans[$i]
            [void]$cboRingPlan.Items.Add(('{0} ({1} rings)' -f $p.Name, @($p.Rings).Count))
            if ($previous -and $p.Name -eq $previous) { $selected = $i }
        }
        if ($script:RingPlans.Count -gt 0) { $cboRingPlan.SelectedIndex = $selected }
        foreach ($e in $script:RingPlanErrors) { Add-LogLine -Message ('Ring plan refused: ' + ($e -replace "`r?`n", ' ')) }
    }
    finally {
        $script:RingSuppress = $false
    }
    Update-RingPlanSelection
}

function Update-RingProgramCombo {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates window controls only.')]
    param()
    $cboRingProgram.Items.Clear()
    $name = ([string]$txtRingObject.Text).Trim()
    if ((Get-RingSelectedType) -ne 'Package' -or -not $name) { return }
    $pkg = Test-PackageExists -PackageName $name
    if (-not $pkg) { return }
    foreach ($p in @(Get-CMPackagePrograms -Package $pkg)) { [void]$cboRingProgram.Items.Add([string]$p.ProgramName) }
    if ($cboRingProgram.Items.Count -eq 1) { $cboRingProgram.SelectedIndex = 0 }
}

function Get-RingPreviewSubmission {
    <#
    .SYNOPSIS
        Commits pending cell edits and parses the edited text back into the
        ring rows. Returns @{ Rows; Errors }.
    #>
    try { [void]$dgRingPreview.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true) } catch { $null = $_ }
    $errors = New-Object System.Collections.Generic.List[string]
    foreach ($row in $script:RingPreviewRows) {
        $label = Get-RingLabel -Ring $row
        $row.CollectionID = ([string]$row.CollectionID).Trim().ToUpperInvariant()
        try { $row.AvailableDateTime = ConvertFrom-RingDateText $row.AvailableText }
        catch { $errors.Add(('{0} available: {1}' -f $label, $_.Exception.Message)) }
        if ($row.Purpose -eq 'Required') {
            try { $row.DeadlineDateTime = ConvertFrom-RingDateText $row.DeadlineText }
            catch { $errors.Add(('{0} deadline: {1}' -f $label, $_.Exception.Message)) }
        }
        else {
            $row.DeadlineDateTime = $null
        }
    }
    return @{ Rows = @($script:RingPreviewRows); Errors = @($errors) }
}

function Format-RingMessageList {
    param([string[]]$Lines, [int]$Max = 12)
    $list = @($Lines)
    $shown = @($list | Select-Object -First $Max)
    $text = '- ' + ($shown -join "`n- ")
    if ($list.Count -gt $Max) { $text += ("`n({0} more in the log)" -f ($list.Count - $Max)) }
    return $text
}

function Invoke-RingValidate {
    Reset-RingValidation
    $type = Get-RingSelectedType
    $name = ([string]$txtRingObject.Text).Trim()
    $plan = Get-RingSelectedPlan
    $program = if ($cboRingProgram.SelectedItem) { [string]$cboRingProgram.SelectedItem } else { '' }

    if (-not $plan) { Add-LogLine -Message 'Select a ring plan.'; Set-StatusText -Text 'No ring plan selected.'; return }
    if (-not $name) { Add-LogLine -Message 'Enter or browse the object to deploy.'; Set-StatusText -Text 'Object is required.'; return }
    if ($type -eq 'Package' -and -not $program) {
        Update-RingProgramCombo
        if ($cboRingProgram.Items.Count -eq 1) { $program = [string]$cboRingProgram.SelectedItem }
        else { Add-LogLine -Message 'Select the package program.'; Set-StatusText -Text 'Program selection required.'; return }
    }

    $sub = Get-RingPreviewSubmission
    if (@($sub.Errors).Count -gt 0) {
        [void](Show-ThemedMessage -Owner $window -Title 'Ring preview' -Message ("Fix the preview dates (yyyy-MM-dd HH:mm):`n" + (Format-RingMessageList $sub.Errors)) -Buttons OK -Icon Error)
        return
    }
    foreach ($row in $sub.Rows) { $row.Checks = '' }

    $exp = Test-RingExpansion -Rings $sub.Rows -Now (Get-RingNow -TimeBasedOn $plan.TimeBasedOn)
    if (@($exp.Errors).Count -gt 0) {
        foreach ($e in $exp.Errors) { Add-LogLine -Message ('Ring rule: ' + $e) }
        Update-RingPreviewGrid
        [void](Show-ThemedMessage -Owner $window -Title 'Ring run blocked' -Message ("The rings break these rules:`n" + (Format-RingMessageList $exp.Errors)) -Buttons OK -Icon Error)
        Set-StatusText -Text 'Ring run blocked by plan rules.'
        return
    }

    if (-not (Connect-IfNeeded)) { return }
    Set-StatusText -Text 'Validating rings...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        $pre = Test-RingPreflight -Type $type -ObjectName $name -ProgramName $program -Rings $sub.Rows
    }
    finally {
        $window.Cursor = $null
    }

    $collections = @{}
    foreach ($res in $pre.Rings) {
        $row = @($sub.Rows | Where-Object { [int]$_.Index -eq $res.Index })[0]
        if (-not $row) { continue }
        $row.Checks = $res.Message
        if ($res.Collection) {
            $row.CollectionName = [string]$res.Collection.Name
            $row.MemberCount    = $res.Collection.MemberCount
            $collections[[int]$res.Index] = $res.Collection
        }
    }
    Update-RingPreviewGrid
    if ($pre.ObjectMessage) { Add-LogLine -Message $pre.ObjectMessage }

    if ($pre.Ok) {
        $script:RingValidated = @{
            Type        = $type
            Name        = $name
            ProgramName = $program
            Object      = $pre.Object
            Plan        = $plan
            Hold        = [bool]$chkRingHold.IsChecked
            Rows        = @($sub.Rows)
            Collections = $collections
            Warnings    = @($exp.Warnings)
        }
        $btnRingCreate.Visibility = [System.Windows.Visibility]::Visible
        Set-StatusText -Text ('All {0} ring(s) passed. Ready to create.' -f @($sub.Rows).Count)
    }
    else {
        Set-StatusText -Text 'Validation complete. Fix the failing rings before creating.'
    }
}

function Invoke-RingOneDeployment {
    <#
    .SYNOPSIS
        Distributes to the ring's DP group when it has one, creates the ring's
        deployment, and writes its audit record. Returns the deployment result.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Validated,
        [Parameter(Mandatory)]$Ring,
        [Parameter(Mandatory)]$Collection,
        [Parameter(Mandatory)][string]$RunId
    )

    $result = $null
    $current = @{} + $Validated
    $identity = Get-RingObjectIdentity -Type $Validated.Type -TargetObject $Validated.Object -ProgramName $Validated.ProgramName
    $pre = Test-RingPreflight -Type $Validated.Type -ObjectName $Validated.Name -ProgramName $Validated.ProgramName `
        -Rings @($Ring) -ExpectedObjectId $identity.ID -SkipContentCheck
    if (-not $pre.Ok) {
        $why = if ($pre.ObjectMessage) { $pre.ObjectMessage } else { $pre.Rings[0].Message }
        $result = @{ Success = $false; DeploymentID = $null; DeploymentUniqueID = $null; Error = ("Pre-create validation failed: {0}" -f $why) }
    }
    else {
        $current.Object = $pre.Object
        $Collection = $pre.Rings[0].Collection

        if ($Ring.DPGroup -and $current.Type -ne 'SUG') {
            $auditGate = Test-DeploymentLogWritable -LogPath $script:DeploymentAuditLog
            if (-not $auditGate.Success) {
                $result = @{ Success = $false; DeploymentID = $null; DeploymentUniqueID = $null; Error = ("Audit log unavailable; content and deployment were not changed: {0}" -f $auditGate.Error) }
            }
        }

        if ($Ring.DPGroup -and $current.Type -ne 'SUG' -and $null -eq $result) {
            $dp = @(Invoke-ContentDistributionToGroups -Type $current.Type -TargetObject $current.Object -DPGroupNames @([string]$Ring.DPGroup))
            $failed = @($dp | Where-Object { -not $_.Success })
            if ($failed.Count -gt 0) {
                $msg = ("Distribution to DP group '{0}' failed: {1}. The deployment was not created." -f $Ring.DPGroup, $failed[0].Error)
                $result = @{ Success = $false; DeploymentID = $null; DeploymentUniqueID = $null; Error = $msg }
            }
            else {
                Add-LogLine -Message ('{0}: content sent to DP group {1}.' -f (Get-RingLabel -Ring $Ring), $Ring.DPGroup)
                # A deployment another session creates during the distribution
                # call would pass the first check; the check runs again here.
                $pre = Test-RingPreflight -Type $current.Type -ObjectName $current.Name -ProgramName $current.ProgramName `
                    -Rings @($Ring) -ExpectedObjectId $identity.ID -SkipContentCheck
                if (-not $pre.Ok) {
                    $why = if ($pre.ObjectMessage) { $pre.ObjectMessage } else { $pre.Rings[0].Message }
                    $result = @{ Success = $false; DeploymentID = $null; DeploymentUniqueID = $null; Error = ("Pre-create validation failed after content distribution: {0}" -f $why) }
                }
                else {
                    $current.Object = $pre.Object
                    $Collection = $pre.Rings[0].Collection
                }
            }
        }

        if ($null -eq $result) {
            $auditGate = Test-DeploymentLogWritable -LogPath $script:DeploymentAuditLog
            if (-not $auditGate.Success) {
                $result = @{ Success = $false; DeploymentID = $null; DeploymentUniqueID = $null; Error = ("Audit log unavailable; deployment was not created: {0}" -f $auditGate.Error) }
            }
            else {
                $result = Invoke-RingDeployment -Type $current.Type -TargetObject $current.Object -Collection $Collection `
                    -Ring $Ring -ProgramName $current.ProgramName -TimeBasedOn $current.Plan.TimeBasedOn
            }
        }
    }
    $record = New-RingAuditRecord -Type $current.Type -TargetObject $current.Object -Collection $Collection -Ring $Ring `
        -Result $result -PlanName $Validated.Plan.Name -RunId $RunId -ProgramName $Validated.ProgramName
    $audit = Write-DeploymentLog -LogPath $script:DeploymentAuditLog -Record $record
    $result.AuditLogged = [bool]$audit.Success
    $result.AuditError = [string]$audit.Error

    if ($result.Success) {
        Add-LogLine -Message ('{0} created: {1} -> {2} (DeploymentID {3}).' -f (Get-RingLabel -Ring $Ring), $Validated.Name, $Collection.Name, $result.DeploymentID)
        if (-not $audit.Success) {
            Add-LogLine -Message ('{0}: deployment exists, but the audit record failed: {1}' -f (Get-RingLabel -Ring $Ring), $audit.Error)
        }
    }
    else {
        if (-not $audit.Success) { $result.Error += (" Audit record also failed: {0}" -f $audit.Error) }
        Add-LogLine -Message ('{0} failed: {1}' -f (Get-RingLabel -Ring $Ring), $result.Error)
    }
    return $result
}

function Get-RingRecordedId {
    param([Parameter(Mandatory)][hashtable]$Result)
    if ($Result.DeploymentUniqueID) { return [string]$Result.DeploymentUniqueID }
    Write-Log ('No unique deployment ID returned; recording {0}. Reconcile confirms it by collection.' -f $Result.DeploymentID) -Level WARN
    return [string]$Result.DeploymentID
}

function Format-RingLine {
    param($Ring, $Collection, [string]$Type, [switch]$Held)
    $when = '{0}, available {1}' -f $Ring.Purpose, (Format-RingDate $Ring.AvailableDateTime)
    if ($Ring.DeadlineDateTime) { $when += ', deadline ' + (Format-RingDate $Ring.DeadlineDateTime) }
    if ($Ring.DPGroup -and $Type -ne 'SUG') { $when += ('; first distribute to DP group {0}' -f $Ring.DPGroup) }
    $heldText = if ($Held) { ' (held)' } else { '' }
    return ("{0}{1} -> {2} ({3}, {4} members)`n    {5}" -f (Get-RingLabel -Ring $Ring), $heldText, $Collection.Name, $Collection.CollectionID, $Collection.MemberCount, $when)
}

function Invoke-RingCreate {
    $v = $script:RingValidated
    if (-not $v) { Add-LogLine -Message 'Run Validate first.'; return }
    $rows = @($v.Rows)

    $lines = foreach ($r in $rows) {
        Format-RingLine -Ring $r -Collection $v.Collections[[int]$r.Index] -Type $v.Type -Held:($v.Hold -and [int]$r.Index -gt 1)
    }
    $count = if ($v.Hold) { 1 } else { $rows.Count }
    $msg = ("Create {0} deployment(s) of {1} '{2}' from plan {3}:`n`n{4}" -f $count, $v.Type, $v.Name, $v.Plan.Name, ($lines -join "`n"))
    if ($v.Plan.TimeBasedOn -eq 'Utc') { $msg += "`n`nAll times are UTC." }
    if ($v.Hold) {
        $msg += ("`n`nHeld rings are created later with Promote. A run file is written to {0}." -f $script:RingRunStatePath)
    }
    if (@($v.Warnings).Count -gt 0) { $msg += ("`n`nWarnings:`n" + (Format-RingMessageList $v.Warnings)) }
    $msg += "`n`nContinue?"
    $icon = if (@($v.Warnings).Count -gt 0) { 'Warn' } else { 'Question' }
    $confirm = Show-ThemedMessage -Owner $window -Title 'Confirm ring deployment' -Message $msg -Buttons OKCancel -Icon $icon
    if ($confirm -ne 'OK') { Add-LogLine -Message 'Ring deployment cancelled by user.'; return }

    $runId = New-RingRunId
    Add-LogLine -Message ('Ring run {0}: plan {1}, {2} ring(s), hold later rings: {3}.' -f $runId, $v.Plan.Name, $rows.Count, $v.Hold)
    Set-StatusText -Text 'Creating ring deployments...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        if ($v.Hold) { Invoke-RingCreateHeld -Validated $v -RunId $runId }
        else         { Invoke-RingCreateAll  -Validated $v -RunId $runId }
    }
    finally {
        $window.Cursor = $null
        Reset-RingValidation
        Update-RingPreviewGrid
    }
    Update-RingRunsGrid -Live:$script:ConnectedToCM
}

function Invoke-RingCreateAll {
    param([Parameter(Mandatory)][hashtable]$Validated, [Parameter(Mandatory)][string]$RunId)
    $rows = @($Validated.Rows)
    $created = @()
    $failure = $null
    foreach ($r in $rows) {
        $res = Invoke-RingOneDeployment -Validated $Validated -Ring $r -Collection $Validated.Collections[[int]$r.Index] -RunId $RunId
        if ($res.Success) {
            $r.Checks = ('Created: {0}' -f $res.DeploymentID)
            $created += $r
            if (-not $res.AuditLogged) {
                $r.Checks += ' (audit write failed)'
                $failure = @{ Ring = $r; Error = [string]$res.AuditError; Created = $true }
                break
            }
        }
        else {
            $r.Checks = ('Failed: {0}' -f $res.Error)
            $failure = @{ Ring = $r; Error = $res.Error }
            break
        }
    }
    if (-not $failure) {
        Set-StatusText -Text ('Created {0} ring deployment(s).' -f $created.Count)
        [void](Show-ThemedMessage -Owner $window -Title 'Ring deployment' -Message ("Created {0} of {0} deployment(s). Run ID {1}. ConfigMgr enforces each ring's dates." -f $created.Count, $RunId) -Buttons OK -Icon Info)
        return
    }
    $notCreated = @($rows | Where-Object { $created -notcontains $_ } | ForEach-Object { Get-RingLabel -Ring $_ })
    Set-StatusText -Text ('Ring run stopped at {0}.' -f (Get-RingLabel -Ring $failure.Ring))
    if ($failure.Created) {
        $text = ("{0} was created, but its audit record could not be written: {1}`n`nLater rings were not created. Created: {2}`nNot created: {3}`n`nThe created deployment stays in Configuration Manager." -f
            (Get-RingLabel -Ring $failure.Ring), $failure.Error,
            (($created | ForEach-Object { Get-RingLabel -Ring $_ }) -join ', '),
            ($notCreated -join ', '))
        [void](Show-ThemedMessage -Owner $window -Title 'Ring run stopped: audit failure' -Message $text -Buttons OK -Icon Warn)
    }
    else {
        $text = ("{0} failed: {1}`n`nCreated: {2}`nNot created: {3}`n`nThe created deployments stay in Configuration Manager." -f
            (Get-RingLabel -Ring $failure.Ring), $failure.Error,
            $(if ($created.Count) { ($created | ForEach-Object { Get-RingLabel -Ring $_ }) -join ', ' } else { 'none' }),
            ($notCreated -join ', '))
        [void](Show-ThemedMessage -Owner $window -Title 'Ring run stopped' -Message $text -Buttons OK -Icon Error)
    }
}

function Invoke-RingCreateHeld {
    param([Parameter(Mandatory)][hashtable]$Validated, [Parameter(Mandatory)][string]$RunId)
    $identity = Get-RingObjectIdentity -Type $Validated.Type -TargetObject $Validated.Object -ProgramName $Validated.ProgramName
    $run = New-RingRun -PlanName $Validated.Plan.Name -Object $identity -Rings @($Validated.Rows) -RunId $RunId -TimeBasedOn $Validated.Plan.TimeBasedOn
    $fileName = Get-RingRunFileName -PlanName $Validated.Plan.Name -ObjectKey $identity.ID -Stamp (Get-Date)

    # The file exists before ring 1 is created, so a crash after the create
    # still leaves a run that Promote can find.
    try { $path = New-RingRunFile -Run $run -Folder $script:RingRunStatePath -FileName $fileName }
    catch {
        [void](Show-ThemedMessage -Owner $window -Title 'Ring run not started' -Message ("The run file could not be written: {0}`n`nNothing was created." -f $_.Exception.Message) -Buttons OK -Icon Error)
        return
    }
    Add-LogLine -Message ('Run file written: {0}' -f $path)

    $lock = $null
    try {
        $lock = Enter-RingRunLock -Path $path
        $ring1 = Get-RingRunRing -Run $run -RingIndex 1
        $row1  = @($Validated.Rows)[0]
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Creating)
        try { Save-RingRun -Run $run -Path $path }
        catch {
            [void](Show-ThemedMessage -Owner $window -Title 'Ring run not started' -Message ("The run could not record its create intent: {0}`n`nNothing was created." -f $_.Exception.Message) -Buttons OK -Icon Error)
            return
        }
        $res = Invoke-RingOneDeployment -Validated $Validated -Ring $ring1 -Collection $Validated.Collections[1] -RunId $RunId
        if (-not $res.Success) {
            [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Held)
            try { Save-RingRun -Run $run -Path $path }
            catch { Add-LogLine -Message ("Ring 1 remains Creating until live reconciliation can confirm no deployment exists: {0}" -f $_.Exception.Message) }
            $row1.Checks = ('Failed: {0}' -f $res.Error)
            [void](Show-ThemedMessage -Owner $window -Title 'Ring 1 failed' -Message ("{0}`n`nThe run stays open. If the run file could not be updated, refresh Open runs after the create recovery wait." -f $res.Error) -Buttons OK -Icon Error)
            return
        }
        $row1.Checks = ('Created: {0}' -f $res.DeploymentID)
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID (Get-RingRecordedId -Result $res) -AssignmentID $res.DeploymentID)
        try { Save-RingRun -Run $run -Path $path }
        catch {
            [void](Show-ThemedMessage -Owner $window -Title 'Run file not updated' -Message ("Ring 1 was created (DeploymentID {0}), but the run file could not be updated: {1}`n`nThe saved run marks ring 1 Creating. Refresh Open runs to reconcile and adopt the deployment; do not retry it manually." -f $res.DeploymentID, $_.Exception.Message) -Buttons OK -Icon Warn)
            return
        }
        Set-StatusText -Text 'Ring 1 created; later rings held.'
        $message = ("Ring 1 created (DeploymentID {0}). {1} later ring(s) are held.`n`nUse Promote in Open runs after you check ring 1." -f $res.DeploymentID, (@($Validated.Rows).Count - 1))
        if (-not $res.AuditLogged) { $message += ("`n`nThe deployment exists, but its audit record failed: {0}" -f $res.AuditError) }
        [void](Show-ThemedMessage -Owner $window -Title 'Ring deployment' -Message $message -Buttons OK -Icon $(if ($res.AuditLogged) { 'Info' } else { 'Warn' }))
    }
    catch {
        [void](Show-ThemedMessage -Owner $window -Title 'Ring deployment' -Message $_.Exception.Message -Buttons OK -Icon Error)
    }
    finally {
        Exit-RingRunLock -LockPath $lock
    }
}

function Update-RingRunsGrid {
    <#
    .SYNOPSIS
        Lists the rings of every open (or closed) run. -Live reads counts from
        the site, reconciles deleted deployments, and closes finished runs.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Writes only run-state files of this tool; the site is read, not changed.')]
    param([switch]$Live)

    $showClosed = [bool]$chkRingShowClosed.IsChecked
    $lblRingRuns.Text = if ($showClosed) { 'CLOSED RUNS' } else { 'OPEN RUNS' }
    $live = $Live -and -not $showClosed
    $gridRows = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]

    foreach ($f in @(Get-RingRunFile -Path $script:RingRunStatePath -Closed:$showClosed)) {
        $run = $null
        $runNote = ''
        $summaries = @{}
        $recoveryNotes = @{}
        $rec = $null
        if ($live) {
            $lock = $null
            try {
                $lock = Enter-RingRunLock -Path $f.FullName
                $run = Read-RingRun -Path $f.FullName
                foreach ($r in @($run.Rings | Where-Object { $_.Status -eq 'Created' })) {
                    $summaries[[int]$r.Index] = Get-RingLiveSummary -Ring $r -ObjectType $run.Object.Type
                }
                $creatingCandidates = @{}
                foreach ($r in @($run.Rings | Where-Object { $_.Status -eq 'Creating' })) {
                    $pre = Test-RingPreflight -Type $run.Object.Type -ObjectName $run.Object.Name -ProgramName $run.Object.ProgramName `
                        -Rings @($r) -ExpectedObjectId $run.Object.ID -SkipContentCheck
                    $ringCheck = $pre.Rings[0]
                    $checkFailed = -not $ringCheck.Collection -or -not $ringCheck.CollectionSafe -or
                        -not [string]::IsNullOrWhiteSpace([string]$pre.ObjectMessage) -or
                        -not [string]::IsNullOrWhiteSpace([string]$ringCheck.DuplicateCheckError)
                    $creatingCandidates[[int]$r.Index] = @{
                        CheckFailed = $checkFailed
                        Error = if ($checkFailed) {
                            if ($pre.ObjectMessage) { [string]$pre.ObjectMessage }
                            elseif (-not $ringCheck.Collection) { [string]$ringCheck.Message }
                            elseif (-not $ringCheck.CollectionSafe) { [string]$ringCheck.Message }
                            else { [string]$ringCheck.DuplicateCheckError }
                        } else { '' }
                        Deployments = @($ringCheck.ExistingDeployments)
                    }
                }
                $rec = Update-RingRunReconcile -Run $run -Summaries $summaries -CreatingCandidates $creatingCandidates
                $recoveryNotes = $rec.RecoveryNotes
                foreach ($i in $rec.Recovered) { Add-LogLine -Message ('Reconcile: {0} ring {1} deployment adopted into the run.' -f $f.Name, $i) }
                foreach ($i in $rec.Reset)     { Add-LogLine -Message ('Reconcile: {0} ring {1} had no matching deployment and returned to Held.' -f $f.Name, $i) }
                foreach ($i in $rec.Removed)   { Add-LogLine -Message ('Reconcile: {0} ring {1} deployment is gone from the site; marked Removed.' -f $f.Name, $i) }
                foreach ($i in $rec.Corrected) { Add-LogLine -Message ('Reconcile: {0} ring {1} deployment ID corrected.' -f $f.Name, $i) }
                foreach ($i in $rec.Recovered) {
                    $recoveredRing = Get-RingRunRing -Run $run -RingIndex ([int]$i)
                    $summaries[[int]$i] = Get-RingLiveSummary -Ring $recoveredRing -ObjectType $run.Object.Type
                }
                if ($rec.Changed) { Save-RingRun -Run $run -Path $f.FullName }
                if ($rec.Finished) {
                    $dest = Close-RingRun -Run $run -Path $f.FullName
                    Add-LogLine -Message ('Reconcile: run finished and moved to {0}.' -f $dest)
                    continue
                }
            }
            catch {
                $runNote = $_.Exception.Message
                if ($rec) {
                    foreach ($i in $rec.Recovered) {
                        $recoveryNotes[[int]$i] = ("A matching deployment was found, but the recovered run state could not be saved: {0}" -f $runNote)
                    }
                    foreach ($i in $rec.Reset) {
                        $recoveryNotes[[int]$i] = ("No matching deployment was found, but the reset to Held could not be saved: {0}" -f $runNote)
                    }
                }
                try { $run = Read-RingRun -Path $f.FullName } catch { $run = $null }
            }
            finally {
                Exit-RingRunLock -LockPath $lock
            }
        }
        else {
            try { $run = Read-RingRun -Path $f.FullName } catch { $runNote = $_.Exception.Message }
        }

        if (-not $run) {
            $gridRows.Add([PSCustomObject]@{
                RunLabel = $f.Name; RunPath = $f.FullName; Index = ''; Name = ''; CollectionID = ''; Status = ''
                DeadlineText = ''; Targeted = ''; Success = ''; Errors = ''; InProgress = ''; Note = $runNote
                ShowPromote = $false; CanPromote = $false; PromoteTip = ''
            })
            continue
        }

        $next = Get-RingRunNextHeld -Run $run
        $finished = Test-RingRunFinished -Run $run
        $label = ('{0}: {1} ({2})' -f $run.PlanName, $run.Object.Name, (Format-RingDate $run.CreatedAt))
        foreach ($r in @($run.Rings | Sort-Object { [int]$_.Index })) {
            $idx = [int]$r.Index
            $s = $summaries[$idx]
            $note = $runNote
            $counted = ($s -and $s.Found -eq $true -and $s.Summarized -eq $true)
            if ($s -and $s.Error) { $note = $s.Error }
            elseif ($s -and $s.Found -eq $true -and -not $counted) { $note = 'The site has not summarized this deployment yet.' }
            elseif ([int]$r.ShiftedMinutes -gt 0) { $note = ('Dates moved {0} minutes later at promote.' -f $r.ShiftedMinutes) }
            elseif ($r.Status -eq 'Created' -and -not $live -and -not $showClosed) { $note = 'Refresh to read live counts.' }
            if ($r.Status -eq 'Creating') {
                $note = if ($recoveryNotes.ContainsKey($idx)) { [string]$recoveryNotes[$idx] } else { 'Create outcome unresolved. Refresh and reconcile before retrying.' }
            }

            $row = [PSCustomObject]@{
                RunLabel     = $label
                RunPath      = $f.FullName
                Index        = $idx
                Name         = [string]$r.Name
                CollectionID = [string]$r.CollectionID
                Status       = [string]$r.Status
                DeadlineText = Format-RingDate $r.DeadlineDateTime
                Targeted     = if ($counted) { $s.Targeted }   else { '' }
                Success      = if ($counted) { $s.Success }    else { '' }
                Errors       = if ($counted) { $s.Errors }     else { '' }
                InProgress   = if ($counted) { $s.InProgress } else { '' }
                Note         = $note
                ShowPromote  = $false
                CanPromote   = $false
                PromoteTip   = ''
            }

            if (-not $showClosed -and -not $finished -and $idx -eq $next -and $r.Status -eq 'Held') {
                $row.ShowPromote = $true
                if ($idx -eq 1) {
                    $row.CanPromote = $true
                    $row.PromoteTip = 'Create ring 1.'
                }
                else {
                    $prev = Get-RingRunRing -Run $run -RingIndex ($idx - 1)
                    $hasThreshold = $null -ne $prev.SuccessThresholdPercent -and ([string]$prev.SuccessThresholdPercent).Trim().Length -gt 0
                    if (-not $hasThreshold) {
                        $row.CanPromote = $true
                        $row.PromoteTip = ('Create ring {0}. Ring {1} has no success threshold.' -f $idx, ($idx - 1))
                    }
                    elseif (-not $live) {
                        $row.PromoteTip = 'Click Refresh and reconcile to read ring counts first.'
                    }
                    else {
                        $t = Test-RingThreshold -Summary $summaries[$idx - 1] -ThresholdPercent $prev.SuccessThresholdPercent
                        $row.CanPromote = [bool]$t.Met
                        $row.PromoteTip = if ($t.Met) { ('Create ring {0}. Ring {1} success is {2}%.' -f $idx, ($idx - 1), $t.Percent) } else { $t.Reason }
                        if (-not $t.Met) { $row.Note = $t.Reason }
                    }
                }
            }
            $gridRows.Add($row)
        }
    }
    $dgRingRuns.ItemsSource = $gridRows
}

function Invoke-RingPromote {
    param([Parameter(Mandatory)][string]$RunPath, [Parameter(Mandatory)][int]$RingIndex)

    if (-not (Connect-IfNeeded)) { return }
    $lock = $null
    try { $lock = Enter-RingRunLock -Path $RunPath }
    catch {
        [void](Show-ThemedMessage -Owner $window -Title 'Promote refused' -Message $_.Exception.Message -Buttons OK -Icon Warn)
        return
    }

    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        # Re-read under the lock: another operator may have promoted or
        # reconciled this run after the grid was drawn.
        $run = Read-RingRun -Path $RunPath
        $check = Test-RingRunPromotable -Run $run -RingIndex $RingIndex
        if (-not $check.Ok) {
            [void](Show-ThemedMessage -Owner $window -Title 'Promote refused' -Message $check.Reason -Buttons OK -Icon Warn)
            return
        }
        $ring = Get-RingRunRing -Run $run -RingIndex $RingIndex

        if ($RingIndex -gt 1) {
            $prev = Get-RingRunRing -Run $run -RingIndex ($RingIndex - 1)
            $prevSummary = Get-RingLiveSummary -Ring $prev -ObjectType $run.Object.Type
            $t = Test-RingThreshold -Summary $prevSummary -ThresholdPercent $prev.SuccessThresholdPercent
            if (-not $t.Met) {
                [void](Show-ThemedMessage -Owner $window -Title 'Promote refused' -Message ('{0}: {1}' -f (Get-RingLabel -Ring $prev), $t.Reason) -Buttons OK -Icon Warn)
                return
            }
        }

        $pre = Test-RingPreflight -Type $run.Object.Type -ObjectName $run.Object.Name -ProgramName $run.Object.ProgramName `
            -Rings @($ring) -ExpectedObjectId $run.Object.ID
        if (-not $pre.Ok) {
            $why = if ($pre.ObjectMessage) { $pre.ObjectMessage } else { $pre.Rings[0].Message }
            [void](Show-ThemedMessage -Owner $window -Title 'Promote refused' -Message ('{0}: {1}' -f (Get-RingLabel -Ring $ring), $why) -Buttons OK -Icon Error)
            return
        }
        $collection = $pre.Rings[0].Collection

        $now = Get-RingNow -TimeBasedOn ([string]$run.TimeBasedOn)
        $shift = Get-RingPromoteShift -Ring $ring -Now $now
        if ($shift -gt 0) { [void](Move-RingRunSchedule -Run $run -FromIndex $RingIndex -Minutes $shift) }
        $exp = Test-RingExpansion -Rings @($ring) -Now $now
        if (@($exp.Errors).Count -gt 0) {
            [void](Show-ThemedMessage -Owner $window -Title 'Promote refused' -Message (Format-RingMessageList $exp.Errors) -Buttons OK -Icon Error)
            return
        }

        $msg = ("Promote ring {0} of plan {1} for {2} '{3}':`n`n{4}" -f $RingIndex, $run.PlanName, $run.Object.Type, $run.Object.Name,
            (Format-RingLine -Ring $ring -Collection $collection -Type $run.Object.Type))
        if ($run.TimeBasedOn -eq 'Utc') { $msg += "`n`nAll times are UTC." }
        if ($shift -gt 0) {
            $msg += ("`n`nThe planned deadline has passed. This ring and every later held ring move {0} minutes later." -f $shift)
        }
        if (@($exp.Warnings).Count -gt 0) { $msg += ("`n`nWarnings:`n" + (Format-RingMessageList $exp.Warnings)) }
        $msg += "`n`nContinue?"
        $confirm = Show-ThemedMessage -Owner $window -Title 'Confirm promote' -Message $msg -Buttons OKCancel -Icon Question
        if ($confirm -ne 'OK') { Add-LogLine -Message 'Promote cancelled by user.'; return }

        $validated = @{
            Type        = [string]$run.Object.Type
            Name        = [string]$run.Object.Name
            ProgramName = [string]$run.Object.ProgramName
            Object      = $pre.Object
            Plan        = [PSCustomObject]@{ Name = [string]$run.PlanName; TimeBasedOn = [string]$run.TimeBasedOn }
        }
        [void](Set-RingRunRingStatus -Run $run -RingIndex $RingIndex -Status Creating)
        try { Save-RingRun -Run $run -Path $RunPath }
        catch {
            [void](Show-ThemedMessage -Owner $window -Title 'Promote not started' -Message ("The run could not record its create intent: {0}`n`nNothing was created." -f $_.Exception.Message) -Buttons OK -Icon Error)
            return
        }
        $res = Invoke-RingOneDeployment -Validated $validated -Ring $ring -Collection $collection -RunId ([string]$run.RunId)
        if (-not $res.Success) {
            [void](Set-RingRunRingStatus -Run $run -RingIndex $RingIndex -Status Held)
            try { Save-RingRun -Run $run -Path $RunPath }
            catch { Add-LogLine -Message ("Ring {0} remains Creating until live reconciliation can confirm no deployment exists: {1}" -f $RingIndex, $_.Exception.Message) }
            [void](Show-ThemedMessage -Owner $window -Title 'Promote failed' -Message ("{0}`n`nIf the run file could not be updated, refresh Open runs after the create recovery wait." -f $res.Error) -Buttons OK -Icon Error)
            return
        }
        [void](Set-RingRunRingStatus -Run $run -RingIndex $RingIndex -Status Created -DeploymentID (Get-RingRecordedId -Result $res) -AssignmentID $res.DeploymentID)
        try { Save-RingRun -Run $run -Path $RunPath }
        catch {
            [void](Show-ThemedMessage -Owner $window -Title 'Run file not updated' -Message ("Ring {0} was created (DeploymentID {1}), but the run file could not be updated: {2}`n`nThe saved run marks this ring Creating. Refresh Open runs to reconcile and adopt the deployment; do not retry it manually." -f $RingIndex, $res.DeploymentID, $_.Exception.Message) -Buttons OK -Icon Warn)
            return
        }
        Set-StatusText -Text ('Ring {0} created.' -f $RingIndex)
        Add-LogLine -Message ('Promote: ring {0} of {1} created.' -f $RingIndex, (Split-Path -Path $RunPath -Leaf))
        if (-not $res.AuditLogged) {
            $auditMessage = ("Ring {0} exists, but its audit record failed: {1}" -f $RingIndex, $res.AuditError)
            Add-LogLine -Message $auditMessage
            Set-StatusText -Text 'Ring created; audit write failed.'
            [void](Show-ThemedMessage -Owner $window -Title 'Audit log write failed' -Message $auditMessage -Buttons OK -Icon Warn)
        }
    }
    catch {
        [void](Show-ThemedMessage -Owner $window -Title 'Promote' -Message $_.Exception.Message -Buttons OK -Icon Error)
    }
    finally {
        $window.Cursor = $null
        Exit-RingRunLock -LockPath $lock
    }
    Update-RingRunsGrid -Live
}

$script:ShowRingView = {
    $script:CurrentType = 'Rings'
    $formPane.Visibility         = [System.Windows.Visibility]::Collapsed
    $ringPane.Visibility         = [System.Windows.Visibility]::Visible
    $pnlApplyTemplate.Visibility = [System.Windows.Visibility]::Collapsed
    $txtModuleHeader.Text    = 'Ring Deployment'
    $txtModuleSubheader.Text = 'Deploy one object to the collections of a ring plan, one ring after another.'
    if (-not $script:RingViewInitialized) {
        $script:RingViewInitialized = $true
        Update-RingPlanCombo
    }
    Update-RingRunsGrid -Live:$script:ConnectedToCM
    Set-StatusText -Text 'Viewing Ring Deployment.'
}

$cboRingType.Add_SelectionChanged({
    $isPackage = ((Get-RingSelectedType) -eq 'Package')
    $vis = if ($isPackage) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $lblRingProgram.Visibility = $vis
    $cboRingProgram.Visibility = $vis
    $cboRingProgram.Items.Clear()
    $txtRingObject.Text = ''
    Reset-RingValidation
})

$btnRingBrowseObject.Add_Click({
    if (-not (Connect-IfNeeded)) { return }
    $type = Get-RingSelectedType
    $b = $script:RingTypeBrowse[$type]
    $picked = Show-BrowseDialog -Owner $window -Title $b.Title -Watermark $b.Watermark -Type $b.Browse -NameProperty $b.NameProperty
    if ($picked) {
        $txtRingObject.Text = $picked
        Add-LogLine -Message ('Ring object selected: {0}' -f $picked)
        if ($type -eq 'Package') { Update-RingProgramCombo }
    }
})

$txtRingObject.Add_TextChanged({
    if ((Get-RingSelectedType) -eq 'Package') { $cboRingProgram.Items.Clear() }
    Reset-RingValidation
})
$cboRingProgram.Add_SelectionChanged({ Reset-RingValidation })
$chkRingHold.Add_Click({ Reset-RingValidation })

$cboRingPlan.Add_SelectionChanged({
    if ($script:RingSuppress) { return }
    Update-RingPlanSelection
})

$btnRingReloadPlans.Add_Click({ Update-RingPlanCombo })

$btnRingOpenFolder.Add_Click({
    if (-not (Test-Path -LiteralPath $script:RingPlanDir)) { New-Item -ItemType Directory -Path $script:RingPlanDir -Force | Out-Null }
    Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $script:RingPlanDir)
})

$btnRingExpand.Add_Click({ Update-RingPreview })

$dgRingPreview.Add_CellEditEnding({ Reset-RingValidation })

$btnRingPickCollection.Add_Click({
    $row = $dgRingPreview.SelectedItem
    if (-not $row) { Add-LogLine -Message 'Select a ring row in the preview first.'; return }
    if (-not (Connect-IfNeeded)) { return }
    $reload = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift) -ne 0
    $picked = Show-CollectionBrowseItem -Owner $window -Force:$reload
    if (-not $picked) { return }
    if (Test-CollectionIdBuiltIn -CollectionId ([string]$picked.CollectionID)) {
        [void](Show-ThemedMessage -Owner $window -Title 'Collection blocked' -Message ('{0} ({1}) is a built-in collection. Rings never target built-in collections.' -f $picked.Name, $picked.CollectionID) -Buttons OK -Icon Error)
        return
    }
    $row.CollectionID   = [string]$picked.CollectionID
    $row.CollectionName = [string]$picked.Name
    $row.Checks         = ''
    Reset-RingValidation
    Update-RingPreviewGrid
    Add-LogLine -Message ('{0}: collection set to {1} ({2}) for this run.' -f (Get-RingLabel -Ring $row), $picked.Name, $picked.CollectionID)
})

$btnRingValidate.Add_Click({ Invoke-RingValidate })
$btnRingCreate.Add_Click({ Invoke-RingCreate })

$btnRingRefreshRuns.Add_Click({
    if ([bool]$chkRingShowClosed.IsChecked) { Update-RingRunsGrid; return }
    if (-not (Connect-IfNeeded)) { Update-RingRunsGrid; return }
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try { Update-RingRunsGrid -Live }
    finally { $window.Cursor = $null }
    Set-StatusText -Text 'Ring runs refreshed.'
})

$chkRingShowClosed.Add_Click({ Update-RingRunsGrid })

$dgRingRuns.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{
    $button = $args[1].OriginalSource
    if ($button -isnot [System.Windows.Controls.Button]) { return }
    $row = $button.DataContext
    if ($null -eq $row -or -not $row.RunPath -or -not $row.ShowPromote) { return }
    Invoke-RingPromote -RunPath ([string]$row.RunPath) -RingIndex ([int]$row.Index)
})

# =============================================================================
# Wire handlers
# =============================================================================
$btnApps.Add_Click({          & $script:SetCurrentType 'Apps' })
$btnPackages.Add_Click({      & $script:SetCurrentType 'Packages' })
$btnTaskSequences.Add_Click({ & $script:SetCurrentType 'TaskSequences' })
$btnSUG.Add_Click({           & $script:SetCurrentType 'SUG' })
$btnRings.Add_Click({         & $script:ShowRingView })

$btnOptions.Add_Click({
    Show-OptionsDialog -Owner $window -InitialSection 'Connection'
})

$radRequired.Add_Checked({
    $dtpDeadline.IsEnabled = $true
    $lblDeadline.Opacity   = 1.0
    # pnlRequiredOptions is Apps-only now. TS/SUG/Packages each own
    # their Required-extras checkboxes inside their own Grid.Row=9
    # panels (chkTs* / chkSug* / chkPkg*). Gating here prevents
    # overlap when CurrentType isn't Apps.
    if ($script:CurrentType -eq 'Apps') {
        $pnlRequiredOptions.Visibility = [System.Windows.Visibility]::Visible
    }
})

$radAvailable.Add_Checked({
    $dtpDeadline.IsEnabled = $false
    $lblDeadline.Opacity   = 0.5
    $pnlRequiredOptions.Visibility = [System.Windows.Visibility]::Collapsed
})

$btnBrowseTarget.Add_Click({
    if (-not (Connect-IfNeeded)) { return }
    switch ($script:CurrentType) {
        'Apps' {
            $picked = Show-BrowseDialog -Owner $window -Title 'Browse applications' `
                -Watermark 'Filter by name, version, or CI ID' `
                -Type 'Apps' -NameProperty 'LocalizedDisplayName'
            if ($picked) {
                $txtTargetName.Text = $picked
                Add-LogLine -Message ('Selected application: {0}' -f $picked)
            }
        }
        'Packages' {
            $picked = Show-BrowseDialog -Owner $window -Title 'Browse packages' `
                -Watermark 'Filter by name, package ID, manufacturer, or version' `
                -Type 'Packages' -NameProperty 'Name'
            if ($picked) {
                $txtTargetName.Text = $picked
                Add-LogLine -Message ('Selected package: {0}' -f $picked)
                $pkg = Test-PackageExists -PackageName $picked
                if ($pkg) { Update-PackageProgramDropdown -Package $pkg }
            }
        }
        'TaskSequences' {
            $picked = Show-BrowseDialog -Owner $window -Title 'Browse task sequences' `
                -Watermark 'Filter by name, package ID, or description' `
                -Type 'TaskSequences' -NameProperty 'Name'
            if ($picked) {
                $txtTargetName.Text = $picked
                Add-LogLine -Message ('Selected task sequence: {0}' -f $picked)
            }
        }
        'SUG' {
            $picked = Show-BrowseDialog -Owner $window -Title 'Browse software update groups' `
                -Watermark 'Filter by name' `
                -Type 'SUG' -NameProperty 'LocalizedDisplayName'
            if ($picked) {
                $txtTargetName.Text = $picked
                Add-LogLine -Message ('Selected SUG: {0}' -f $picked)
            }
        }
        default {
            Add-LogLine -Message ('{0} browse dialog lands in a later session.' -f $script:CurrentType)
        }
    }
})

$btnBrowseCollection.Add_Click({
    if (-not (Connect-IfNeeded)) { return }
    $reload = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift) -ne 0
    $picked = Show-CollectionBrowse -Owner $window -Force:$reload
    if ($picked) {
        $txtCollection.Text = $picked
        Add-LogLine -Message ('Selected collection: {0}' -f $picked)
    }
})

$txtTargetName.Add_TextChanged({
    # Packages: clear program dropdown when target changes (stale program list after a name edit)
    if ($script:CurrentType -eq 'Packages') { $cboProgram.Items.Clear() }
    if ($script:AllChecksPassed) { Reset-ValidationUi }
})
$txtCollection.Add_TextChanged({ if ($script:AllChecksPassed) { Reset-ValidationUi } })
$cboProgram.Add_SelectionChanged({ if ($script:AllChecksPassed) { Reset-ValidationUi } })

$btnDPPicker.Add_Click({
    if (-not (Connect-IfNeeded)) { return }
    $target = Get-CurrentDistributionTarget
    if (-not $target) {
        Add-LogLine -Message ('Enter a valid {0} name before picking DP groups.' -f $script:CurrentType)
        return
    }
    $groups = Get-DPGroupsLoaded
    if (-not $groups -or @($groups).Count -eq 0) {
        Add-LogLine -Message 'No DP groups found in this site.'
        return
    }
    # WMI content-tracking identifier differs by type. See module docstring.
    # Apps: SMS_DPGroupContentInfo keys off the ModelName (no /version suffix),
    # but Get-CMApplication's CI_UniqueID includes a trailing "/N" revision.
    # Strip the version so the lookup matches all revisions of the same app.
    $contentId = if ($script:CurrentType -eq 'Apps' -and $target.PSObject.Properties['CI_UniqueID'] -and $target.CI_UniqueID) {
        ([string]$target.CI_UniqueID) -replace '/\d+$', ''
    } else {
        [string]$target.PackageID
    }
    $alreadyTargeted = @(Get-ContentTargetedDPGroups -ObjectID $contentId)
    $picked = Show-DPGroupPickerDialog -Owner $window -AllGroupNames $groups `
        -PreSelected $global:SelectedDPGroups `
        -AlreadyTargeted $alreadyTargeted
    if ($null -ne $picked) {
        $global:SelectedDPGroups = @($picked)
        Update-DPPickerButtonLabel
        if (@($picked).Count -gt 0) {
            Add-LogLine -Message ('DP groups selected: {0}' -f ($picked -join ', '))
        }
    }
})

$btnDistributeContent.Add_Click({
    if (@($global:SelectedDPGroups).Count -eq 0) {
        Add-LogLine -Message 'Select one or more DP groups first.'
        return
    }
    if (-not (Connect-IfNeeded)) { return }
    $target = Get-CurrentDistributionTarget
    if (-not $target) {
        Add-LogLine -Message ('Enter a valid {0} name before distributing.' -f $script:CurrentType)
        return
    }
    $type = Get-CurrentDistributionType
    if (-not $type) { return }

    Set-StatusText -Text 'Distributing content...'
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try {
        $results = Invoke-ContentDistributionToGroups -Type $type -TargetObject $target -DPGroupNames $global:SelectedDPGroups
        $ok = @($results | Where-Object { $_.Success }).Count
        $fail = @($results | Where-Object { -not $_.Success }).Count
        Add-LogLine -Message ('Distribution submitted: {0} succeeded, {1} failed.' -f $ok, $fail)
        if ($fail -eq 0) {
            Set-StatusText -Text ('Distribution submitted for {0} DP group(s).' -f $ok)
        } else {
            Set-StatusText -Text ('Distribution completed with {0} failure(s). See log.' -f $fail)
        }
        # Invalidate previous validation so the user re-checks with fresh distribution status
        if ($script:AllChecksPassed) { Reset-ValidationUi }
    }
    finally {
        $window.Cursor = $null
    }
})

$btnValidate.Add_Click({
    switch ($script:CurrentType) {
        'Apps'          { & $script:InvokeAppsValidate }
        'Packages'      { & $script:InvokePackagesValidate }
        'TaskSequences' { & $script:InvokeTaskSequencesValidate }
        'SUG'           { & $script:InvokeSUGValidate }
    }
})

$btnDeploy.Add_Click({
    switch ($script:CurrentType) {
        'Apps'          { & $script:InvokeAppsDeploy }
        'Packages'      { & $script:InvokePackagesDeploy }
        'TaskSequences' { & $script:InvokeTaskSequencesDeploy }
        'SUG'           { & $script:InvokeSUGDeploy }
    }
})

# =============================================================================
# Log drawer polish: context menu (Clear / Copy all)
# =============================================================================
$logContextMenu = New-Object System.Windows.Controls.ContextMenu
$miClearLog = New-Object System.Windows.Controls.MenuItem
$miClearLog.Header = 'Clear log'
$miCopyLog  = New-Object System.Windows.Controls.MenuItem
$miCopyLog.Header  = 'Copy log to clipboard'
[void]$logContextMenu.Items.Add($miClearLog)
[void]$logContextMenu.Items.Add($miCopyLog)
$txtLog.ContextMenu = $logContextMenu
$miClearLog.Add_Click({ $txtLog.Clear() })
$miCopyLog.Add_Click({
    if (-not [string]::IsNullOrEmpty($txtLog.Text)) {
        [System.Windows.Clipboard]::SetText($txtLog.Text)
        Set-StatusText -Text 'Log copied to clipboard.'
    }
})

# =============================================================================
# Apply Template combo (main window header)
#
# Combo always has "(Select a template...)" at index 0 as a sentinel so the
# user can pick any template by selecting it from the list. Changing selection
# off the sentinel prefills the form fields from the template. Selecting the
# sentinel is a no-op. Repopulating is idempotent -- the Templates panel calls
# $script:RefreshApplyTemplateCombo after every New / Duplicate / Save /
# Delete so the combo stays in sync without a full app restart.
# =============================================================================
$script:ApplyTemplateSuppress = $false
$script:ApplyTemplateTemplatesDir = Join-Path $PSScriptRoot 'Templates'

$script:RefreshApplyTemplateCombo = {
    $script:ApplyTemplateSuppress = $true
    try {
        $cboApplyTemplate.Items.Clear()
        $sentinel = New-Object System.Windows.Controls.ComboBoxItem
        $sentinel.Content = '(Select a template...)'
        $sentinel.Tag = $null
        [void]$cboApplyTemplate.Items.Add($sentinel)

        if (Test-Path -LiteralPath $script:ApplyTemplateTemplatesDir) {
            $files = @(Get-ChildItem -LiteralPath $script:ApplyTemplateTemplatesDir -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object Name)
            foreach ($f in $files) {
                try {
                    $data = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json
                    $name = if ($data.Name) { [string]$data.Name } else { [System.IO.Path]::GetFileNameWithoutExtension($f.Name) }
                    $item = New-Object System.Windows.Controls.ComboBoxItem
                    $item.Content = $name
                    $item.Tag     = $data
                    [void]$cboApplyTemplate.Items.Add($item)
                } catch { Write-Log ('Apply-template combo skipping ' + $f.Name + ': ' + $_) -Level WARN }
            }
        }
        $cboApplyTemplate.SelectedIndex = 0
    } finally {
        $script:ApplyTemplateSuppress = $false
    }
}

# Map notification tag -> index in the main form's $cboNotification (matches
# MainWindow.xaml item order).
$script:NotificationTagIndex = @{
    'DisplayAll'                  = 0
    'DisplaySoftwareCenterOnly'   = 1
    'HideAll'                     = 2
}

$script:ApplyTemplateToForm = {
    param($data)
    # Collection (all four types share $txtCollection)
    if ($data.TargetCollectionName) { $txtCollection.Text = [string]$data.TargetCollectionName }

    # Purpose
    if (([string]$data.DeployPurpose) -eq 'Required') {
        $radRequired.IsChecked = $true
    } else {
        $radAvailable.IsChecked = $true
    }

    # Notification (Apps only -- template still carries it, but main form
    # only exposes it for Apps; safe to set regardless since the control
    # exists and is just hidden for other types)
    $notifTag = [string]$data.UserNotification
    if ($notifTag -and $script:NotificationTagIndex.ContainsKey($notifTag)) {
        $cboNotification.SelectedIndex = $script:NotificationTagIndex[$notifTag]
    }

    # Time basis
    if (([string]$data.TimeBasedOn) -eq 'UTC') {
        $radUtc.IsChecked = $true
    } else {
        $radLocalTime.IsChecked = $true
    }

    # Required-only options. Four parallel sets of checkboxes live on
    # four type-scoped panels (Apps -> pnlRequiredOptions; Packages ->
    # pnlPackageOptions; SUG -> pnlSUGOptions; TS -> pnlTaskSequenceOptions).
    # Set all four from the same template fields so switching Type
    # preserves intent.
    $chkOverrideSW.IsChecked      = [bool]$data.OverrideServiceWindow
    $chkRebootOutSW.IsChecked     = [bool]$data.RebootOutsideServiceWindow
    $chkMetered.IsChecked         = [bool]$data.AllowMeteredConnection
    $chkPkgOverrideSW.IsChecked   = [bool]$data.OverrideServiceWindow
    $chkPkgRebootOutSW.IsChecked  = [bool]$data.RebootOutsideServiceWindow
    $chkPkgMetered.IsChecked      = [bool]$data.AllowMeteredConnection
    $chkSugOverrideSW.IsChecked   = [bool]$data.OverrideServiceWindow
    $chkSugAllowRestart.IsChecked = [bool]$data.RebootOutsideServiceWindow
    $chkSugMetered.IsChecked      = [bool]$data.AllowMeteredConnection
    $chkTsOverrideSW.IsChecked    = [bool]$data.OverrideServiceWindow
    $chkTsRebootOutSW.IsChecked   = [bool]$data.RebootOutsideServiceWindow
    $chkTsMetered.IsChecked       = [bool]$data.AllowMeteredConnection

    # SUG options (checkbox names on the main form)
    $chkBoundaryFallback.IsChecked = if ($null -ne $data.AllowBoundaryFallback) { [bool]$data.AllowBoundaryFallback } else { $true }
    $chkMSFallback.IsChecked       = [bool]$data.AllowMicrosoftUpdate
    $chkFullScan.IsChecked         = if ($null -ne $data.RequirePostRebootFullScan) { [bool]$data.RequirePostRebootFullScan } else { $true }

    # Deadline offset -- only meaningful when Required. If Required and offset > 0,
    # compute Deadline from current Available.
    if ($data.DeployPurpose -eq 'Required' -and $null -ne $data.DefaultDeadlineOffsetHours) {
        $hours = [int]$data.DefaultDeadlineOffsetHours
        if ($hours -gt 0) {
            $base = $dtpAvailable.SelectedDateTime
            if ($null -eq $base) { $base = Get-Date }
            $dtpDeadline.SelectedDateTime = $base.AddHours($hours)
        }
    }

    Add-LogLine -Message ('Template applied: {0}' -f [string]$data.Name)
}

$cboApplyTemplate.Add_SelectionChanged({
    if ($script:ApplyTemplateSuppress) { return }
    if ($cboApplyTemplate.SelectedIndex -le 0) { return }
    $item = $cboApplyTemplate.SelectedItem
    if ($null -eq $item -or $null -eq $item.Tag) { return }
    & $script:ApplyTemplateToForm $item.Tag
})

# =============================================================================
# Window lifecycle
# =============================================================================
$script:SavedDarkTheme = $true
$script:SavedType      = 'Apps'
Restore-WindowState -Window $window -Path (Get-WindowStatePath) -OnStateLoaded {
    param($s)
    $script:SavedDarkTheme = if ($null -ne $s.DarkTheme) { [bool]$s.DarkTheme } else { $true }
    $script:SavedType      = if ($s.CurrentType)         { [string]$s.CurrentType } else { 'Apps' }
}

$window.Add_SourceInitialized({
    if (-not $script:SavedDarkTheme) {
        $toggleTheme.IsOn = $false
    }
    else {
        Set-ButtonTheme -IsDark $true
    }
})

$window.Add_Loaded({
    if ($script:SavedType -eq 'Rings') {
        & $script:SetCurrentType 'Apps'
        & $script:ShowRingView
    }
    elseif ($script:TypeMeta.ContainsKey($script:SavedType)) {
        & $script:SetCurrentType $script:SavedType
    }
    else {
        & $script:SetCurrentType 'Apps'
    }
    & $script:RefreshApplyTemplateCombo
    Add-LogLine -Message ('Deployment Helper loaded. Site={0} Provider={1}' -f $global:Prefs['SiteCode'], $global:Prefs['SMSProvider'])
    Add-LogLine -Message ('Tool log: {0}' -f $toolLogPath)
    Set-StatusText -Text 'Ready.'
})

$window.Add_Closing({
    Save-WindowState -Window $window -Path (Get-WindowStatePath) -ExtraState @{
        DarkTheme   = ($toggleTheme.IsOn -eq $true)
        CurrentType = $script:CurrentType
    }
    Stop-BgWork
    Close-SuiteBgRunspace -Runspace $script:BgRunspace
    $script:BgRunspace = $null
    try { Disconnect-CMSite } catch { }
})

# =============================================================================
# Show window
# =============================================================================
# Unhandled-exception capture. Normal transcripts don't flush during a
# process-terminating crash, so write directly to a dedicated file from
# both handlers. Two places an unhandled exception can come from:
#   - WPF dispatcher (UI-thread event handlers: SourceInitialized, Loaded,
#     button Click, etc.)
#   - AppDomain (background threads, finalizers, P/Invoke callbacks)
$global:__crashLog = Join-Path $__txDir ('DeploymentHelper-crash-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

$__writeCrash = {
    param($Source, $Exception)
    try {
        $lines = @()
        $lines += ('=== ' + $Source + ' @ ' + (Get-Date -Format 'o') + ' ===')
        $lines += ('Type   : ' + $Exception.GetType().FullName)
        $lines += ('Message: ' + $Exception.Message)
        $lines += ('Stack  :')
        $lines += ([string]$Exception.StackTrace).Split([Environment]::NewLine)
        $inner = $Exception.InnerException
        $depth = 1
        while ($inner) {
            $lines += ('--- InnerException depth ' + $depth + ' ---')
            $lines += ('Type   : ' + $inner.GetType().FullName)
            $lines += ('Message: ' + $inner.Message)
            $lines += ('Stack  :')
            $lines += ([string]$inner.StackTrace).Split([Environment]::NewLine)
            $inner = $inner.InnerException
            $depth++
        }
        [System.IO.File]::AppendAllText($global:__crashLog, (($lines -join [Environment]::NewLine) + [Environment]::NewLine))
    } catch { }
}
$global:__writeCrash = $__writeCrash

# WPF dispatcher event uses Add_UnhandledException in PS event syntax.
$window.Dispatcher.Add_UnhandledException({
    param($s, $e)
    & $global:__writeCrash 'DispatcherUnhandledException' $e.Exception
    $e.Handled = $false
})

# AppDomain for non-UI-thread exceptions.
[AppDomain]::CurrentDomain.Add_UnhandledException({
    param($s, $e)
    & $global:__writeCrash 'AppDomainUnhandledException' ([Exception]$e.ExceptionObject)
})

[void]$window.ShowDialog()
