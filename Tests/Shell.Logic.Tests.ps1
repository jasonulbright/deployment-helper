#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Logic of start-deploymenthelper.ps1. The script cannot be dot-sourced (it opens
# the window), so each function is cut out of its syntax tree and defined here,
# with plain objects in place of the window controls. The commands of the module
# and of SuiteCommon that these functions call are mocked, so no site is needed.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase -ErrorAction SilentlyContinue
    foreach ($dll in 'Microsoft.Xaml.Behaviors.dll', 'ControlzEx.dll', 'MahApps.Metro.dll') { [void][System.Reflection.Assembly]::LoadFrom((Join-Path $script:AppRoot "Lib\$dll")) }
    Import-AppModule
    $script:ShellPath = Join-Path $script:AppRoot 'start-deploymenthelper.ps1'
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
    $wanted = 'Get-DhPreferences', 'Save-DhPreferences', 'Get-WindowStatePath', 'Add-LogLine', 'Set-StatusText',
        'Show-AuditWriteFailure', 'Test-DeploymentAuditReady', 'Test-NoCurrentDuplicateDeployment',
        'Update-FooterSite', 'Set-CheckGlyph', 'Reset-ValidationUi', 'Set-NotificationItems', 'Connect-IfNeeded',
        'Test-DeploymentNotification', 'Test-DeploymentSchedule',
        'Get-DPGroupsLoaded', 'Update-DPPickerButtonLabel', 'Get-CurrentDistributionTarget', 'Get-CurrentDistributionType',
        'Get-RingSelectedType', 'Get-RingSelectedPlan', 'Format-RingDate', 'Reset-RingValidation',
        'Format-RingMessageList', 'Format-RingLine', 'Get-RingRecordedId', 'Format-RolloutPlanLine'
    foreach ($n in $wanted) {
        $def = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true)
        if (-not $def) { throw "function $n not found in the script" }
        # The functions read $PSScriptRoot, which is empty in a scriptblock that was cut out, so it points to the app folder here.
        . ([scriptblock]::Create($def.Extent.Text.Replace('$PSScriptRoot', '$script:AppRoot')))
    }
}

Describe 'Test-DeploymentNotification' {
    It 'blocks the purpose Available with the notification Hide All' {
        $r = Test-DeploymentNotification -Purpose 'Available' -Notification 'HideAll'
        $r.Ok | Should -BeFalse
        $r.Level | Should -Be 'Error'
        $r.Reason | Should -Match 'incompatible'
    }

    It 'allows the purpose <Purpose> with the notification <Notification>' -ForEach @(
        @{ Purpose = 'Available'; Notification = 'DisplayAll' }
        @{ Purpose = 'Available'; Notification = 'DisplaySoftwareCenterOnly' }
        @{ Purpose = 'Required'; Notification = 'DisplayAll' }
        @{ Purpose = 'Required'; Notification = 'DisplaySoftwareCenterOnly' }
        @{ Purpose = 'Required'; Notification = 'HideAll' }
    ) {
        $r = Test-DeploymentNotification -Purpose $Purpose -Notification $Notification
        $r.Ok | Should -BeTrue
        $r.Level | Should -Be 'Info'
        $r.Reason | Should -Be ''
    }

    It 'refuses a notification that it does not know' {
        { Test-DeploymentNotification -Purpose 'Available' -Notification 'Loud' } | Should -Throw
    }
}

Describe 'Test-DeploymentSchedule' {
    BeforeAll { $script:Now = [datetime]'2026-11-16T12:00:00' }

    It 'accepts a Required deployment with a deadline after the available time and after now' {
        $r = Test-DeploymentSchedule -Available $script:Now.AddHours(1) -Deadline $script:Now.AddHours(2) -Purpose Required -Now $script:Now
        $r.Ok | Should -BeTrue
        $r.Level | Should -Be 'Info'
    }

    It 'blocks a Required deployment without a deadline' {
        $r = Test-DeploymentSchedule -Available $script:Now.AddHours(1) -Deadline $null -Purpose Required -Now $script:Now
        $r.Ok | Should -BeFalse
        $r.Level | Should -Be 'Error'
        $r.Reason | Should -Be 'Deadline is required when Purpose is Required.'
    }

    It 'blocks a deadline on or before the available time' -ForEach @(@{ Hours = 1 }, @{ Hours = 0 }) {
        $avail = $script:Now.AddHours(2)
        $r = Test-DeploymentSchedule -Available $avail -Deadline $avail.AddHours(-$Hours) -Purpose Required -Now $script:Now
        $r.Ok | Should -BeFalse
        $r.Reason | Should -Match 'must be after Available'
    }

    It 'blocks a deadline in the past and names it' {
        $r = Test-DeploymentSchedule -Available $script:Now.AddHours(-3) -Deadline $script:Now.AddHours(-1) -Purpose Required -Now $script:Now
        $r.Ok | Should -BeFalse
        $r.Reason | Should -Match 'is in the past'
        $r.Reason | Should -Match '2026-11-16 11:00'
    }

    It 'does not need a deadline for the purpose Available, and ignores one that is not valid' {
        (Test-DeploymentSchedule -Available $script:Now.AddHours(1) -Deadline $null -Purpose Available -Now $script:Now).Level | Should -Be 'Info'
        (Test-DeploymentSchedule -Available $script:Now.AddHours(1) -Deadline $script:Now.AddHours(-9) -Purpose Available -Now $script:Now).Ok | Should -BeTrue
    }

    It 'warns, and does not block, when the available time is more than five minutes in the past' {
        $r = Test-DeploymentSchedule -Available $script:Now.AddMinutes(-6) -Deadline $null -Purpose Available -Now $script:Now
        $r.Ok | Should -BeTrue
        $r.Level | Should -Be 'Warn'
        $r.Reason | Should -Match 'is in the past'
    }

    It 'treats an available time within five minutes of now as on time' {
        (Test-DeploymentSchedule -Available $script:Now.AddMinutes(-5) -Deadline $null -Purpose Available -Now $script:Now).Level | Should -Be 'Info'
    }

    It 'warns when the available time is more than five years away' {
        $r = Test-DeploymentSchedule -Available $script:Now.AddYears(5).AddDays(1) -Deadline $null -Purpose Available -Now $script:Now
        $r.Ok | Should -BeTrue
        $r.Level | Should -Be 'Warn'
        $r.Reason | Should -Match 'five years'
        (Test-DeploymentSchedule -Available $script:Now.AddYears(5) -Deadline $null -Purpose Available -Now $script:Now).Level | Should -Be 'Info'
    }
}

Describe 'Update-DPPickerButtonLabel' {
    BeforeEach {
        $btnDPPicker = [PSCustomObject]@{ Content = $null }
        $global:SelectedDPGroups = @()
        $global:DPGroupsCache = $null
    }
    AfterAll {
        Remove-Variable -Name SelectedDPGroups, DPGroupsCache -Scope Global -ErrorAction SilentlyContinue
    }

    It 'asks for a selection when no group is selected' {
        Update-DPPickerButtonLabel
        $btnDPPicker.Content | Should -Be 'Select DP Groups...'
    }

    It 'says that all groups are selected' {
        $global:DPGroupsCache = @('A', 'B', 'C')
        $global:SelectedDPGroups = @('A', 'B', 'C')
        Update-DPPickerButtonLabel
        $btnDPPicker.Content | Should -Be 'All 3 DP groups selected'
    }

    It 'counts the selected groups of the loaded ones' {
        $global:DPGroupsCache = @('A', 'B', 'C')
        $global:SelectedDPGroups = @('A')
        Update-DPPickerButtonLabel
        $btnDPPicker.Content | Should -Be '1 of 3 DP groups selected'
    }

    It 'never shows more selected groups than the total when the list is not loaded' {
        $global:SelectedDPGroups = @('A', 'B')
        Update-DPPickerButtonLabel
        $btnDPPicker.Content | Should -Be '2 of 2 DP groups selected'
    }
}

Describe 'Distribution target of the current workflow' {
    BeforeEach {
        $txtTargetName = [PSCustomObject]@{ Text = '  Target One  ' }
        Mock Test-ApplicationExists { [PSCustomObject]@{ Kind = 'app'; Name = $ApplicationName } }
        Mock Test-PackageExists { [PSCustomObject]@{ Kind = 'pkg'; Name = $PackageName } }
        Mock Test-TaskSequenceExists { [PSCustomObject]@{ Kind = 'ts'; Name = $TaskSequenceName } }
    }

    It 'maps the workflow <Type> to the distribution type <Expected>' -ForEach @(
        @{ Type = 'Apps'; Expected = 'Application' }
        @{ Type = 'Packages'; Expected = 'Package' }
        @{ Type = 'TaskSequences'; Expected = 'TaskSequence' }
        @{ Type = 'SUG'; Expected = $null }
        @{ Type = 'Rings'; Expected = $null }
    ) {
        $script:CurrentType = $Type
        Get-CurrentDistributionType | Should -Be $Expected
    }

    It 'looks up the trimmed name with the lookup of the workflow <Type>' -ForEach @(
        @{ Type = 'Apps'; Kind = 'app' }
        @{ Type = 'Packages'; Kind = 'pkg' }
        @{ Type = 'TaskSequences'; Kind = 'ts' }
    ) {
        $script:CurrentType = $Type
        $t = Get-CurrentDistributionTarget
        $t.Kind | Should -Be $Kind
        $t.Name | Should -Be 'Target One'
    }

    It 'returns nothing for an empty name, and calls no lookup' {
        $script:CurrentType = 'Apps'
        $txtTargetName.Text = '   '
        Get-CurrentDistributionTarget | Should -BeNullOrEmpty
        $txtTargetName.Text = $null
        Get-CurrentDistributionTarget | Should -BeNullOrEmpty
        Should -Invoke Test-ApplicationExists -Times 0 -Exactly
    }

    It 'returns nothing for a workflow without a distribution target' {
        $script:CurrentType = 'SUG'
        Get-CurrentDistributionTarget | Should -BeNullOrEmpty
    }
}

Describe 'Log line and status text' {
    BeforeEach {
        $txtStatus = [PSCustomObject]@{ Text = '' }
        $txtLog = [PSCustomObject]@{ Text = '' }
        $txtLog | Add-Member -MemberType ScriptMethod -Name AppendText -Value { param($t) $this.Text += $t }
        $txtLog | Add-Member -MemberType ScriptMethod -Name ScrollToEnd -Value { $script:Scrolled++ }
        $script:Scrolled = 0
    }

    It 'sets the status text' {
        Set-StatusText -Text 'Ready.'
        $txtStatus.Text | Should -Be 'Ready.'
    }

    It 'starts the log with the first line and appends the next ones on new lines' {
        Add-LogLine -Message 'first'
        $txtLog.Text | Should -Match '^\d\d:\d\d:\d\d  first$'
        Add-LogLine -Message 'second'
        $lines = $txtLog.Text -split "`r?`n"
        $lines.Count | Should -Be 2
        $lines[1] | Should -Match '^\d\d:\d\d:\d\d  second$'
        $script:Scrolled | Should -Be 2
    }

    It 'refuses an empty message' {
        { Add-LogLine -Message '' } | Should -Throw
    }
}

Describe 'Audit log guards' {
    BeforeEach {
        $txtStatus = [PSCustomObject]@{ Text = '' }
        $txtLog = [PSCustomObject]@{ Text = '' }
        $txtLog | Add-Member -MemberType ScriptMethod -Name AppendText -Value { param($t) $this.Text += $t }
        $txtLog | Add-Member -MemberType ScriptMethod -Name ScrollToEnd -Value { }
        $window = [PSCustomObject]@{ Name = 'owner' }
        $script:DeploymentAuditLog = 'C:\audit\deployment-audit.jsonl'
        $script:Shown = [System.Collections.ArrayList]::new()
        Mock Show-ThemedMessage { [void]$script:Shown.Add(@{ Title = $Title; Message = $Message; Icon = $Icon }); 'OK' }
    }

    It 'lets a deployment go on when the audit log is writable' {
        Mock Test-DeploymentLogWritable { @{ Success = $true; Error = $null } }
        Test-DeploymentAuditReady | Should -BeTrue
        $script:Shown.Count | Should -Be 0
    }

    It 'blocks a deployment, tells the user and logs it when the audit log is not writable' {
        Mock Test-DeploymentLogWritable { @{ Success = $false; Error = 'Access denied.' } }
        Test-DeploymentAuditReady | Should -BeFalse
        $txtStatus.Text | Should -Be 'Deployment blocked: audit log unavailable.'
        $txtLog.Text | Should -Match 'not writable at C:\\audit\\deployment-audit\.jsonl: Access denied\.'
        $script:Shown.Count | Should -Be 1
        $script:Shown[0].Title | Should -Be 'Deployment Blocked'
        $script:Shown[0].Icon | Should -Be 'Error'
    }

    It 'names the deployment ID when the deployment exists but the audit write failed' {
        Show-AuditWriteFailure -AuditResult @{ Error = 'Disk full.' } -DeploymentCreated $true -DeploymentID 'MCM2001A'
        $txtStatus.Text | Should -Be 'Deployment created; audit write failed.'
        $script:Shown[0].Message | Should -Match 'created deployment MCM2001A'
        $script:Shown[0].Message | Should -Match 'Disk full\.'
    }

    It 'says that nothing was created when the attempt could not be written' {
        Show-AuditWriteFailure -AuditResult @{ Error = 'Disk full.' } -DeploymentCreated $false
        $txtStatus.Text | Should -Be 'Audit write failed.'
        $script:Shown[0].Message | Should -Match '^The deployment attempt could not be written'
    }
}

Describe 'Test-NoCurrentDuplicateDeployment' {
    BeforeEach {
        $txtStatus = [PSCustomObject]@{ Text = '' }
        $txtLog = [PSCustomObject]@{ Text = '' }
        $txtLog | Add-Member -MemberType ScriptMethod -Name AppendText -Value { param($t) $this.Text += $t }
        $txtLog | Add-Member -MemberType ScriptMethod -Name ScrollToEnd -Value { }
        $window = [PSCustomObject]@{ Name = 'owner' }
        $script:Shown = [System.Collections.ArrayList]::new()
        Mock Show-ThemedMessage { [void]$script:Shown.Add($Message); 'OK' }
        $script:Target = [PSCustomObject]@{ LocalizedDisplayName = 'App One'; PackageID = 'MCM00010' }
        $script:Coll = [PSCustomObject]@{ Name = 'Pilot' }
    }

    It 'calls the check of the type <Type> and passes when it finds no deployment' -ForEach @(
        @{ Type = 'Application'; Command = 'Test-DuplicateDeployment' }
        @{ Type = 'Package'; Command = 'Test-DuplicatePackageDeployment' }
        @{ Type = 'TaskSequence'; Command = 'Test-DuplicateTaskSequenceDeployment' }
        @{ Type = 'SUG'; Command = 'Test-DuplicateSUGDeployment' }
    ) {
        Mock Test-DuplicateDeployment { $null }
        Mock Test-DuplicatePackageDeployment { $null }
        Mock Test-DuplicateTaskSequenceDeployment { $null }
        Mock Test-DuplicateSUGDeployment { $null }
        Test-NoCurrentDuplicateDeployment -Type $Type -TargetObject $script:Target -Collection $script:Coll -ProgramName 'Install' | Should -BeTrue
        Should -Invoke $Command -Times 1 -Exactly
        $script:Shown.Count | Should -Be 0
    }

    It 'blocks when a matching deployment now exists' {
        Mock Test-DuplicateDeployment { [PSCustomObject]@{ DeploymentID = 'MCM2001A' } }
        Test-NoCurrentDuplicateDeployment -Type Application -TargetObject $script:Target -Collection $script:Coll | Should -BeFalse
        $script:Shown[0] | Should -Be "A matching deployment now exists on 'Pilot'. No duplicate was created."
        $txtStatus.Text | Should -Be 'Deployment blocked by the final duplicate check.'
    }

    It 'blocks, and creates nothing, when the recheck itself failed' {
        Mock Test-DuplicateDeployment { [PSCustomObject]@{ DuplicateCheckFailed = $true; Error = 'The site did not answer.' } }
        Test-NoCurrentDuplicateDeployment -Type Application -TargetObject $script:Target -Collection $script:Coll | Should -BeFalse
        $script:Shown[0] | Should -Be 'Could not recheck for a duplicate deployment: The site did not answer.. No deployment was created.'
    }
}

Describe 'Validation glyphs' {
    BeforeEach {
        $script:GlyphPending = [char]0x22EF
        $script:GlyphPass = [char]0x2713
        $script:GlyphFail = [char]0x2717
        1..5 | ForEach-Object { Set-Variable -Name "glyphCheck$_" -Value ([PSCustomObject]@{ Text = '' }) }
    }

    It 'shows the glyph of the state <State> by shape, and never sets a color' -ForEach @(
        @{ State = 'Pending'; Code = 0x22EF }
        @{ State = 'Pass'; Code = 0x2713 }
        @{ State = 'Fail'; Code = 0x2717 }
    ) {
        Set-CheckGlyph -Index 3 -State $State
        [int][char]$glyphCheck3.Text | Should -Be $Code
        $glyphCheck3.PSObject.Properties.Name | Should -Not -Contain 'Foreground'
        $glyphCheck1.Text | Should -Be ''
    }

    It 'refuses an index outside 1 to 5 and a state that it does not know' {
        { Set-CheckGlyph -Index 0 -State Pass } | Should -Throw
        { Set-CheckGlyph -Index 6 -State Pass } | Should -Throw
        { Set-CheckGlyph -Index 1 -State 'Done' } | Should -Throw
    }

    It 'resets all five glyphs, forgets every validated object and hides the Deploy button' {
        1..5 | ForEach-Object { Set-CheckGlyph -Index $_ -State Pass }
        $script:ValidatedApp = 'a'; $script:ValidatedCollection = 'c'; $script:ValidatedPackage = 'p'
        $script:ValidatedProgram = 'g'; $script:ValidatedTaskSequence = 't'; $script:ValidatedSUG = 's'
        $script:AllChecksPassed = $true
        $btnDeploy = [PSCustomObject]@{ Visibility = [System.Windows.Visibility]::Visible }
        Reset-ValidationUi
        foreach ($i in 1..5) { [int][char](Get-Variable "glyphCheck$i").Value.Text | Should -Be 0x22EF }
        $script:ValidatedApp | Should -BeNullOrEmpty
        $script:ValidatedCollection | Should -BeNullOrEmpty
        $script:ValidatedPackage | Should -BeNullOrEmpty
        $script:ValidatedProgram | Should -BeNullOrEmpty
        $script:ValidatedTaskSequence | Should -BeNullOrEmpty
        $script:ValidatedSUG | Should -BeNullOrEmpty
        $script:AllChecksPassed | Should -BeFalse
        [string]$btnDeploy.Visibility | Should -Be 'Collapsed'
    }
}

Describe 'Set-NotificationItems' {
    BeforeEach {
        $cboNotification = [PSCustomObject]@{ Items = [System.Collections.ArrayList]::new(); SelectedIndex = -1 }
        $lblNotification = [PSCustomObject]@{ Text = '' }
    }

    It 'fills the box with one item for each entry, selects the first and sets the label' {
        Set-NotificationItems -Items @(@{ Content = 'One'; Tag = 't1' }, @{ Content = 'Two'; Tag = 't2' }) -LabelText 'Notification:'
        $cboNotification.Items.Count | Should -Be 2
        $cboNotification.Items[1].Content | Should -Be 'Two'
        $cboNotification.Items[1].Tag | Should -Be 't2'
        $cboNotification.SelectedIndex | Should -Be 0
        $lblNotification.Text | Should -Be 'Notification:'
    }

    It 'replaces the old items' {
        Set-NotificationItems -Items @(@{ Content = 'One'; Tag = 't1' }) -LabelText 'A'
        Set-NotificationItems -Items @(@{ Content = 'Two'; Tag = 't2' }) -LabelText 'B'
        $cboNotification.Items.Count | Should -Be 1
        $cboNotification.Items[0].Tag | Should -Be 't2'
    }

    It 'selects nothing for an empty list' {
        Set-NotificationItems -Items @() -LabelText 'None'
        $cboNotification.Items.Count | Should -Be 0
        $cboNotification.SelectedIndex | Should -Be -1
    }

    It 'offers the notification sets that the script defines' {
        $sets = @{}
        foreach ($n in 'AppsNotificationItems', 'TaskSequenceAvailabilityItems') {
            $a = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.AssignmentStatementAst] -and $x.Left.Extent.Text -eq ('$script:' + $n) }, $true)
            $a | Should -Not -BeNullOrEmpty
            $sets[$n] = @([regex]::Matches($a.Right.Extent.Text, "Tag = '([^']+)'") | ForEach-Object { $_.Groups[1].Value })
        }
        $sets['AppsNotificationItems'] | Should -Be @('DisplayAll', 'DisplaySoftwareCenterOnly', 'HideAll')
        $sets['TaskSequenceAvailabilityItems'] | Should -Be @('Clients', 'ClientsMediaAndPxe', 'MediaAndPxe', 'MediaAndPxeHidden')
    }
}

Describe 'Connect-IfNeeded' {
    BeforeEach {
        $txtStatus = [PSCustomObject]@{ Text = '' }
        $txtLog = [PSCustomObject]@{ Text = '' }
        $txtLog | Add-Member -MemberType ScriptMethod -Name AppendText -Value { param($t) $this.Text += $t }
        $txtLog | Add-Member -MemberType ScriptMethod -Name ScrollToEnd -Value { }
        $window = [PSCustomObject]@{ Cursor = $null }
        $global:Prefs = @{ SiteCode = 'MCM'; SMSProvider = 'cm01.contoso.com' }
        $script:ConnectedToCM = $false
        Mock Update-FooterSite { }
        Mock Test-CMConnection { $true }
        Mock Connect-CMSite { $true }
    }
    AfterAll { Remove-Variable -Name Prefs -Scope Global -ErrorAction SilentlyContinue }

    It 'does not connect again when the session is connected' {
        $script:ConnectedToCM = $true
        Connect-IfNeeded | Should -BeTrue
        Should -Invoke Connect-CMSite -Times 0 -Exactly
    }

    It 'connects again when the flag is set but the connection is gone' {
        $script:ConnectedToCM = $true
        Mock Test-CMConnection { $false }
        Connect-IfNeeded | Should -BeTrue
        Should -Invoke Connect-CMSite -Times 1 -Exactly
    }

    It 'says what to do when the <Case> is missing, and calls no site' -ForEach @(
        @{ Case = 'site code'; Site = ''; Provider = 'cm01' }
        @{ Case = 'provider'; Site = 'MCM'; Provider = '  ' }
    ) {
        $global:Prefs = @{ SiteCode = $Site; SMSProvider = $Provider }
        Connect-IfNeeded | Should -BeFalse
        $txtStatus.Text | Should -Be 'Connection not configured.'
        $txtLog.Text | Should -Match 'Open Options > Connection'
        Should -Invoke Connect-CMSite -Times 0 -Exactly
    }

    It 'connects with the saved site code and provider, resets the cursor and updates the footer' {
        Connect-IfNeeded | Should -BeTrue
        Should -Invoke Connect-CMSite -Times 1 -Exactly -ParameterFilter { $SiteCode -eq 'MCM' -and $SMSProvider -eq 'cm01.contoso.com' }
        Should -Invoke Update-FooterSite -Times 1 -Exactly
        $script:ConnectedToCM | Should -BeTrue
        $txtStatus.Text | Should -Be 'Connected to MCM.'
        $window.Cursor | Should -BeNullOrEmpty
    }

    It 'reports a failed connection in plain text and stays not connected' {
        Mock Connect-CMSite { $false }
        Connect-IfNeeded | Should -BeFalse
        $script:ConnectedToCM | Should -BeFalse
        $txtStatus.Text | Should -Be 'Connection failed.'
        $txtLog.Text | Should -Match 'Failed to connect\. Check site code, provider, and ConfigurationManager module\.'
    }

    It 'resets the cursor when the connect call throws' {
        Mock Connect-CMSite { throw 'boom' }
        { Connect-IfNeeded } | Should -Throw
        $window.Cursor | Should -BeNullOrEmpty
    }
}

Describe 'Get-DPGroupsLoaded' {
    BeforeEach {
        $txtStatus = [PSCustomObject]@{ Text = '' }
        $txtLog = [PSCustomObject]@{ Text = '' }
        $txtLog | Add-Member -MemberType ScriptMethod -Name AppendText -Value { param($t) $this.Text += $t }
        $txtLog | Add-Member -MemberType ScriptMethod -Name ScrollToEnd -Value { }
        $global:DPGroupsCache = $null
        Mock Connect-IfNeeded { $true }
        Mock Get-DPGroupList { @([PSCustomObject]@{ Name = 'DPG1' }, [PSCustomObject]@{ Name = 'DPG2' }) }
    }
    AfterAll { Remove-Variable -Name DPGroupsCache -Scope Global -ErrorAction SilentlyContinue }

    It 'loads the group names once and keeps them in the cache' {
        Get-DPGroupsLoaded | Should -Be @('DPG1', 'DPG2')
        Get-DPGroupsLoaded | Should -Be @('DPG1', 'DPG2')
        Should -Invoke Get-DPGroupList -Times 1 -Exactly
        $txtLog.Text | Should -Match 'Loaded 2 distribution point group\(s\)\.'
    }

    It 'loads again on a forced reload' {
        $null = Get-DPGroupsLoaded
        $null = Get-DPGroupsLoaded -ForceReload
        Should -Invoke Get-DPGroupList -Times 2 -Exactly
    }

    It 'returns an empty list and reads no group when the site is not connected' {
        Mock Connect-IfNeeded { $false }
        @(Get-DPGroupsLoaded).Count | Should -Be 0
        Should -Invoke Get-DPGroupList -Times 0 -Exactly
        $global:DPGroupsCache | Should -BeNullOrEmpty
    }
}

Describe 'Ring helpers' {
    It 'formats a ring date as text of the form yyyy-MM-dd HH:mm' -ForEach @(
        @{ Value = [datetime]'2026-11-16T09:05:59'; Expected = '2026-11-16 09:05' }
        @{ Value = '2026-11-16T09:05:59'; Expected = '2026-11-16 09:05' }
        @{ Value = '2026-11-16 09:05'; Expected = '2026-11-16 09:05' }
        @{ Value = $null; Expected = '' }
        @{ Value = '   '; Expected = '' }
    ) {
        Format-RingDate -Value $Value | Should -Be $Expected
    }

    It 'returns the text that it cannot parse as it is' {
        Format-RingDate -Value 'next Tuesday' | Should -Be 'next Tuesday'
    }

    It 'reads the selected ring type from the tag, and defaults to Application' {
        $cboRingType = [PSCustomObject]@{ SelectedItem = $null }
        Get-RingSelectedType | Should -Be 'Application'
        $cboRingType.SelectedItem = [PSCustomObject]@{ Tag = 'SUG' }
        Get-RingSelectedType | Should -Be 'SUG'
    }

    It 'returns the plan at the selected index, and nothing outside the list' -ForEach @(
        @{ Index = 0; Expected = 'first' }
        @{ Index = 1; Expected = 'second' }
        @{ Index = 2; Expected = $null }
        @{ Index = -1; Expected = $null }
    ) {
        $script:RingPlans = @([PSCustomObject]@{ Name = 'first' }, [PSCustomObject]@{ Name = 'second' })
        $cboRingPlan = [PSCustomObject]@{ SelectedIndex = $Index }
        (Get-RingSelectedPlan).Name | Should -Be $Expected
    }

    It 'clears the validated plan and hides the Create button' {
        $script:RingValidated = 'something'
        $btnRingCreate = [PSCustomObject]@{ Visibility = [System.Windows.Visibility]::Visible }
        Reset-RingValidation
        $script:RingValidated | Should -BeNullOrEmpty
        [string]$btnRingCreate.Visibility | Should -Be 'Collapsed'
    }

    It 'lists the messages with a dash and counts what is left out' {
        Format-RingMessageList -Lines @('a', 'b') | Should -Be "- a`n- b"
        $text = Format-RingMessageList -Lines (1..15 | ForEach-Object { "m$_" })
        ($text -split "`n").Count | Should -Be 13
        $text | Should -Match '\(3 more in the log\)$'
        Format-RingMessageList -Lines @('a', 'b', 'c') -Max 2 | Should -Match '\(1 more in the log\)$'
    }

    It 'describes a ring line with its collection, purpose, dates and DP group' {
        $ring = [PSCustomObject]@{ Index = 2; Name = 'Pilot'; Purpose = 'Required'; AvailableDateTime = [datetime]'2026-11-16T08:00:00'; DeadlineDateTime = [datetime]'2026-11-18T17:30:00'; DPGroup = 'DPG1' }
        $coll = [PSCustomObject]@{ Name = 'Pilot devices'; CollectionID = 'MCM00042'; MemberCount = 17 }
        $line = Format-RingLine -Ring $ring -Collection $coll -Type 'Application'
        $line | Should -Be "Ring 2 'Pilot' -> Pilot devices (MCM00042, 17 members)`n    Required, available 2026-11-16 08:00, deadline 2026-11-18 17:30; first distribute to distribution point group DPG1"
    }

    It 'leaves out the deadline for an Available ring, the DP group for a software update group, and marks a held ring' {
        $ring = [PSCustomObject]@{ Index = 1; Name = 'Early'; Purpose = 'Available'; AvailableDateTime = [datetime]'2026-11-16T08:00:00'; DeadlineDateTime = $null; DPGroup = 'DPG1' }
        $coll = [PSCustomObject]@{ Name = 'Early'; CollectionID = 'MCM00001'; MemberCount = 3 }
        $line = Format-RingLine -Ring $ring -Collection $coll -Type 'SUG' -Held
        $line | Should -Match "^Ring 1 'Early' \(held\) -> Early"
        $line | Should -Not -Match 'deadline'
        $line | Should -Not -Match 'DP group'
    }

    It 'records the unique deployment ID when there is one, and the deployment ID with a warning when there is not' {
        Mock Write-Log { }
        Get-RingRecordedId -Result @{ DeploymentUniqueID = '{GUID-1}'; DeploymentID = 'MCM2001A' } | Should -Be '{GUID-1}'
        Should -Invoke Write-Log -Times 0 -Exactly
        Get-RingRecordedId -Result @{ DeploymentUniqueID = $null; DeploymentID = 'MCM2001A' } | Should -Be 'MCM2001A'
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter { $Level -eq 'WARN' -and $Message -match 'MCM2001A' }
    }
}

Describe 'Preferences file' {
    BeforeEach { $global:PrefsPath = Join-Path $TestDrive ('prefs-{0}.json' -f [guid]::NewGuid().ToString('N')) }
    AfterAll { Remove-Variable -Name PrefsPath -Scope Global -ErrorAction SilentlyContinue }

    It 'gives the defaults on first run, with no stored path (empty means the default place)' {
        $p = Get-DhPreferences
        $p.SiteCode | Should -Be ''
        $p.SMSProvider | Should -Be ''
        $p.DeploymentAuditLogPath | Should -Be ''
        $p.RingRunStatePath | Should -Be ''
        $p.RolloutRunPath | Should -Be ''
    }

    It 'treats a saved old default under the tool folder as empty, and keeps a path that the user chose' {
        $json = [ordered]@{
            SiteCode               = 'AB1'
            DeploymentAuditLogPath = (Join-Path $script:AppRoot 'Logs\deployment-audit.jsonl')
            RingRunStatePath       = '\\files.example.com\share\runs'
            RolloutRunPath         = ((Join-Path $script:AppRoot 'Rollouts') + '\')
        } | ConvertTo-Json
        Set-Content -LiteralPath $global:PrefsPath -Value $json
        $p = Get-DhPreferences
        $p.SiteCode | Should -Be 'AB1'
        $p.DeploymentAuditLogPath | Should -Be ''
        $p.RingRunStatePath | Should -Be '\\files.example.com\share\runs'
        $p.RolloutRunPath | Should -Be ''
    }

    It 'gives the defaults for a <Case> file' -ForEach @(
        @{ Case = 'corrupt'; Content = '{ not json' }
        @{ Case = 'empty'; Content = '' }
        @{ Case = 'array'; Content = '[1,2,3]' }
    ) {
        Set-Content -LiteralPath $global:PrefsPath -Value $Content
        $p = Get-DhPreferences
        $p.SiteCode | Should -Be ''
        $p.DeploymentAuditLogPath | Should -Be ''
    }

    It 'ignores a key that it does not know, and a blank value' {
        Set-Content -LiteralPath $global:PrefsPath -Value '{ "SiteCode": "MCM", "SMSProvider": "   ", "Evil": "x" }'
        $p = Get-DhPreferences
        $p.SiteCode | Should -Be 'MCM'
        $p.SMSProvider | Should -Be ''
        $p.ContainsKey('Evil') | Should -BeFalse
    }

    It 'round-trips the value <Case>' -ForEach @(
        @{ Case = 'plain'; Value = 'MCM' }
        @{ Case = 'with quotes and a backslash'; Value = 'a "q" \ b' }
        @{ Case = 'of 5000 characters'; Value = ('x' * 5000) }
    ) {
        Save-DhPreferences -Prefs @{ SiteCode = $Value; SMSProvider = 'cm01'; DeploymentAuditLogPath = 'C:\x\audit.jsonl'; RingRunStatePath = 'C:\x\runs' }
        $p = Get-DhPreferences
        $p.SiteCode | Should -Be $Value
        $p.SMSProvider | Should -Be 'cm01'
        $p.DeploymentAuditLogPath | Should -Be 'C:\x\audit.jsonl'
    }

    It 'round-trips non-ASCII text' {
        $value = 'caf' + [char]0x00E9 + '-' + [char]0x65E5 + [char]0x672C
        Save-DhPreferences -Prefs @{ SiteCode = $value }
        (Get-DhPreferences).SiteCode | Should -Be $value
    }

    It 'does not throw when the file cannot be written' {
        $global:PrefsPath = Join-Path $TestDrive 'dir-in-the-way'
        New-Item -ItemType Directory -Path $global:PrefsPath | Out-Null
        { Save-DhPreferences -Prefs @{ SiteCode = 'X' } } | Should -Not -Throw
    }

    It 'keeps the window state file beside the script' {
        Get-WindowStatePath | Should -BeLike '*\DeploymentHelper.windowstate.json'
    }
}

Describe 'Format-RolloutPlanLine' {
    BeforeAll { $script:Sched = @{ Available = [datetime]'2026-11-21T08:00:00'; Deadline = [datetime]'2026-11-24T08:00:00' } }

    It 'says that a new collection is created, with the members and the times' {
        $item = [PSCustomObject]@{ Application = '7-Zip'; Version = '26.06'; CollectionName = '7-Zip - Server - Installed - Version LT 26.06'; Exists = $false; MemberCount = 12 }
        $text = Format-RolloutPlanLine -Item $item -Schedule $script:Sched
        $text | Should -Match '7-Zip 26.06 -> 7-Zip - Server - Installed - Version LT 26.06'
        $text | Should -Match 'create the collection, 12 member\(s\)'
        $text | Should -Match 'available 2026-11-21 08:00, deadline 2026-11-24 08:00'
    }

    It 'says that an existing collection is reused and that uncounted members are not counted' {
        $item = [PSCustomObject]@{ Application = 'A'; Version = '2'; CollectionName = 'Test PCs'; Exists = $true; MemberCount = $null }
        $text = Format-RolloutPlanLine -Item $item -Schedule $script:Sched
        $text | Should -Match 'reuse the existing collection, members not counted'
    }
}
