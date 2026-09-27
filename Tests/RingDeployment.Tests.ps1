#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }

<#
.SYNOPSIS
    Pester 5 tests for ring plans, ring runs, and ring deployment in the
    DeploymentHelperCommon module.

.DESCRIPTION
    Unit tests with mocked ConfigurationManager cmdlets.
    Run: Invoke-Pester -Path .\Tests\RingDeployment.Tests.ps1 -Output Detailed
#>

BeforeAll {
    . (Join-Path $PSScriptRoot 'CMStubs.ps1')
    Set-CMTestStub -Name 'Get-CMCollection', 'Get-CMDeployment', 'New-CMSchedule', 'New-CMApplicationDeployment',
        'New-CMSoftwareUpdateDeployment', 'New-CMPackageDeployment', 'New-CMTaskSequenceDeployment',
        'Get-CMApplicationDeployment', 'Get-CMPackageDeployment', 'Get-CMTaskSequenceDeployment', 'Get-CMUpdateGroupDeployment',
        'Get-CMApplication', 'Get-CMPackage', 'Get-CMTaskSequence', 'Get-CMSoftwareUpdateGroup'

    Import-Module (Join-Path $PSScriptRoot '..\Module\DeploymentHelperCommon.psd1') -Force
    Initialize-Logging -LogPath (Join-Path $TestDrive 'rings.log')

    function New-TestRing {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Test helper; builds test data.')]
        param(
            [int]$Index, [string]$Name, [string]$CollectionID = '', [string]$Purpose = 'Required',
            $Available, $Deadline, [string]$UserNotification = 'DisplayAll'
        )
        [PSCustomObject]@{
            Index = $Index; Name = $Name; CollectionID = $CollectionID; CollectionName = ''
            Purpose = $Purpose; UserNotification = $UserNotification
            AvailableDateTime = $Available; DeadlineDateTime = $Deadline
            OverrideServiceWindow = $false; RebootOutsideServiceWindow = $false; AllowMeteredConnection = $false
            AllowBoundaryFallback = $true; AllowMicrosoftUpdate = $false; RequirePostRebootFullScan = $true
            ShowTaskSequenceProgress = $true; TaskSequenceAvailability = 'Clients'
            FastNetworkOption = 'DownloadContentFromDistributionPointAndRunLocally'
            SlowNetworkOption = 'DoNotRunProgram'; RerunBehavior = 'NeverRerunDeployedProgram'
            DPGroup = ''; SuccessThresholdPercent = $null
        }
    }

    function Write-TestPlan {
        param([string]$Path, [hashtable]$Plan)
        $Plan | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Path -Encoding UTF8
    }

    function New-TestPlanHashtable {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Test helper; builds test data.')]
        param([array]$Rings, [string]$Name = 'Test-Plan')
        @{ Name = $Name; TimeBasedOn = 'LocalTime'; HoldLaterRings = $false; Rings = $Rings }
    }

    function New-TestRunFile {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Test helper; builds test data.')]
        param([string]$Folder, [int]$RingCount = 3)
        $rings = for ($i = 1; $i -le $RingCount; $i++) {
            New-TestRing -Index $i -Name ("R$i") -CollectionID ('MCM0010{0}' -f $i) `
                -Available ([datetime]'2026-10-01 08:00').AddDays(($i - 1) * 7) `
                -Deadline ([datetime]'2026-10-02 08:00').AddDays(($i - 1) * 7)
        }
        $object = [ordered]@{ Type = 'Application'; ID = 'MCM00099'; Name = '7-Zip'; ProgramName = $null }
        $run = New-RingRun -PlanName 'Workstation-Rings' -Object $object -Rings @($rings) -RunId 'run-1' -CreatedAt ([datetime]'2026-09-30 10:00')
        $path = Join-Path $Folder (Get-RingRunFileName -PlanName 'Workstation-Rings' -ObjectKey 'MCM00099' -Stamp ([datetime]'2026-09-30 10:00'))
        Save-RingRun -Run $run -Path $path
        return $path
    }
}

Describe 'ConvertFrom-RingDateText' {
    It 'Parses <Text>' -TestCases @(
        @{ Text = '2026-10-01T08:30:00' }, @{ Text = '2026-10-01 08:30:00' }, @{ Text = '2026-10-01 08:30' }, @{ Text = '2026-10-01T08:30' }
    ) {
        ConvertFrom-RingDateText $Text | Should -Be ([datetime]'2026-10-01 08:30')
    }

    It 'Returns null for empty input' {
        ConvertFrom-RingDateText '' | Should -BeNullOrEmpty
        ConvertFrom-RingDateText $null | Should -BeNullOrEmpty
    }

    It 'Throws on a date in another form' {
        { ConvertFrom-RingDateText '01/10/2026 08:30' } | Should -Throw '*yyyy-MM-dd HH:mm*'
    }

    It 'Round-trips through ConvertTo-RingDateText' {
        ConvertTo-RingDateText ([datetime]'2026-10-01 08:30') | Should -Be '2026-10-01T08:30:00'
    }
}

Describe 'Ring plan seeds' {
    It 'Seeds pass plan validation and name no collection' {
        foreach ($seed in Get-RingPlanSeed) {
            $result = ConvertTo-RingPlan -InputObject $seed
            $result.Errors | Should -BeNullOrEmpty
            @($result.Plan.Rings | Where-Object { $_.CollectionID }).Count | Should -Be 0
        }
    }

    It 'Seeds Workstation-Rings with QA, Pilot, Prod 1, Prod Final and Server-Rings with Test, Prod' {
        $seeds = Get-RingPlanSeed
        ($seeds | Where-Object { $_.Name -eq 'Workstation-Rings' }).Rings.Name | Should -Be @('QA', 'Pilot', 'Prod 1', 'Prod Final')
        ($seeds | Where-Object { $_.Name -eq 'Server-Rings' }).Rings.Name | Should -Be @('Test', 'Prod')
    }

    It 'Writes seeds into an empty folder once and never overwrites' {
        $dir = Join-Path $TestDrive 'seed'
        Initialize-RingPlanFolder -Path $dir | Should -Be 2
        $file = Join-Path $dir 'Workstation-Rings.json'
        Set-Content -LiteralPath $file -Value '{"Name":"Edited","Rings":[]}' -Encoding UTF8
        Initialize-RingPlanFolder -Path $dir | Should -Be 0
        (Get-Content -LiteralPath $file -Raw) | Should -Match 'Edited'
    }

    It 'A seed plan cannot run: every ring has no target' {
        $dir = Join-Path $TestDrive 'seed-run'
        [void](Initialize-RingPlanFolder -Path $dir)
        $plan = Import-RingPlan -Path (Join-Path $dir 'Server-Rings.json')
        $rows = Expand-RingPlan -Plan $plan -Start ([datetime]'2026-10-01 08:00')
        $v = Test-RingExpansion -Rings $rows -Now ([datetime]'2026-09-30')
        @($v.Errors).Count | Should -Be 2
        $v.Errors[0] | Should -Match "Ring 1 'Test' has no target collection"
    }
}

Describe 'Import-RingPlan' {
    BeforeAll {
        $script:PlanDir = Join-Path $TestDrive 'plans'
        New-Item -ItemType Directory -Path $script:PlanDir -Force | Out-Null
        function Get-GoodRing { param($Name, $Id, $A, $D) @{ Name = $Name; CollectionID = $Id; Purpose = 'Required'; AvailableOffsetDays = $A; DeadlineOffsetDays = $D } }
    }

    It 'Loads a valid plan and fills option defaults' {
        $path = Join-Path $script:PlanDir 'good.json'
        Write-TestPlan -Path $path -Plan (New-TestPlanHashtable -Rings @(
            (Get-GoodRing -Name 'QA' -Id 'MCM00101' -A 0 -D 1), (Get-GoodRing -Name 'Pilot' -Id 'mcm00102' -A 2 -D 5)))
        $plan = Import-RingPlan -Path $path
        @($plan.Rings).Count | Should -Be 2
        $plan.Rings[1].CollectionID | Should -Be 'MCM00102'
        $plan.Rings[0].UserNotification | Should -Be 'DisplayAll'
        $plan.Rings[0].AllowBoundaryFallback | Should -BeTrue
        $plan.FilePath | Should -Be $path
    }

    It 'Refuses a plan that names <Id> and calls out the ring' -TestCases @(@{ Id = 'SMS00001' }, @{ Id = 'SMSDM003' }) {
        $path = Join-Path $script:PlanDir 'builtin.json'
        Write-TestPlan -Path $path -Plan (New-TestPlanHashtable -Rings @(
            (Get-GoodRing -Name 'QA' -Id 'MCM00101' -A 0 -D 1), (Get-GoodRing -Name 'Everyone' -Id $Id -A 2 -D 5)))
        { Import-RingPlan -Path $path } | Should -Throw "*Ring 2 'Everyone' targets built-in collection $Id*"
    }

    It 'Refuses rings that share a collection' {
        $path = Join-Path $script:PlanDir 'shared.json'
        Write-TestPlan -Path $path -Plan (New-TestPlanHashtable -Rings @(
            (Get-GoodRing -Name 'QA' -Id 'MCM00101' -A 0 -D 1), (Get-GoodRing -Name 'Pilot' -Id 'MCM00101' -A 2 -D 5)))
        { Import-RingPlan -Path $path } | Should -Throw '*must not share a collection*'
    }

    It 'Refuses a deadline offset that is not later than the previous ring' {
        $path = Join-Path $script:PlanDir 'order.json'
        Write-TestPlan -Path $path -Plan (New-TestPlanHashtable -Rings @(
            (Get-GoodRing -Name 'QA' -Id 'MCM00101' -A 0 -D 5), (Get-GoodRing -Name 'Pilot' -Id 'MCM00102' -A 1 -D 5)))
        { Import-RingPlan -Path $path } | Should -Throw "*Ring 2 'Pilot' deadline offset (5) must be later than the Ring 1 'QA' deadline offset (5)*"
    }

    It 'Refuses a deadline offset on or before the available offset' {
        $path = Join-Path $script:PlanDir 'inverted.json'
        Write-TestPlan -Path $path -Plan (New-TestPlanHashtable -Rings @((Get-GoodRing -Name 'QA' -Id 'MCM00101' -A 3 -D 3)))
        { Import-RingPlan -Path $path } | Should -Throw '*must be later than its available offset*'
    }

    It 'Ignores Available rings when ordering deadlines' {
        $path = Join-Path $script:PlanDir 'mixed.json'
        Write-TestPlan -Path $path -Plan (New-TestPlanHashtable -Rings @(
            (Get-GoodRing -Name 'QA' -Id 'MCM00101' -A 0 -D 2),
            @{ Name = 'Opt-in'; CollectionID = 'MCM00102'; Purpose = 'Available'; AvailableOffsetDays = 1 },
            (Get-GoodRing -Name 'Prod' -Id 'MCM00103' -A 3 -D 6)))
        $plan = Import-RingPlan -Path $path
        $plan.Rings[1].DeadlineOffsetDays | Should -BeNullOrEmpty
    }

    It 'Refuses Available with HideAll' {
        $path = Join-Path $script:PlanDir 'hidden.json'
        Write-TestPlan -Path $path -Plan (New-TestPlanHashtable -Rings @(
            @{ Name = 'QA'; CollectionID = 'MCM00101'; Purpose = 'Available'; AvailableOffsetDays = 0; UserNotification = 'HideAll' }))
        { Import-RingPlan -Path $path } | Should -Throw '*never installs*'
    }

    It 'Refuses a threshold outside 0 to 100 and a malformed collection ID' {
        $path = Join-Path $script:PlanDir 'bad-values.json'
        $r = Get-GoodRing -Name 'QA' -Id 'MCM1' -A 0 -D 1
        $r.SuccessThresholdPercent = 150
        Write-TestPlan -Path $path -Plan (New-TestPlanHashtable -Rings @($r))
        $err = { Import-RingPlan -Path $path } | Should -Throw -PassThru
        $err.Exception.Message | Should -Match 'SuccessThresholdPercent'
        $err.Exception.Message | Should -Match "not an 8-character collection ID"
    }

    It 'Refuses a file that is not JSON' {
        $path = Join-Path $script:PlanDir 'broken.json'
        Set-Content -LiteralPath $path -Value '{ not json' -Encoding UTF8
        { Import-RingPlan -Path $path } | Should -Throw '*broken.json failed to load*'
    }

    It 'Get-RingPlanList returns good plans and the errors of refused ones' {
        $dir = Join-Path $TestDrive 'list'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-TestPlan -Path (Join-Path $dir 'a.json') -Plan (New-TestPlanHashtable -Name 'A' -Rings @((Get-GoodRing -Name 'QA' -Id 'MCM00101' -A 0 -D 1)))
        Write-TestPlan -Path (Join-Path $dir 'b.json') -Plan (New-TestPlanHashtable -Name 'B' -Rings @((Get-GoodRing -Name 'QA' -Id 'SMS00001' -A 0 -D 1)))
        $list = Get-RingPlanList -Path $dir
        @($list.Plans).Count | Should -Be 1
        $list.Plans[0].Name | Should -Be 'A'
        @($list.Errors).Count | Should -Be 1
        $list.Errors[0] | Should -Match 'b.json'
    }
}

Describe 'Expand-RingPlan' {
    It 'Computes absolute dates from day offsets' {
        $plan = (ConvertTo-RingPlan -InputObject (New-TestPlanHashtable -Rings @(
            @{ Name = 'QA'; CollectionID = 'MCM00101'; Purpose = 'Required'; AvailableOffsetDays = 0; DeadlineOffsetDays = 1.5 },
            @{ Name = 'Opt-in'; CollectionID = 'MCM00102'; Purpose = 'Available'; AvailableOffsetDays = 3 }))).Plan
        $rows = Expand-RingPlan -Plan $plan -Start ([datetime]'2026-10-01 08:00')
        $rows[0].AvailableDateTime | Should -Be ([datetime]'2026-10-01 08:00')
        $rows[0].DeadlineDateTime  | Should -Be ([datetime]'2026-10-02 20:00')
        $rows[1].AvailableDateTime | Should -Be ([datetime]'2026-10-04 08:00')
        $rows[1].DeadlineDateTime  | Should -BeNullOrEmpty
        $rows[1].Index | Should -Be 2
    }

    It 'Returns a new row set; the plan is unchanged' {
        $plan = (ConvertTo-RingPlan -InputObject (New-TestPlanHashtable -Rings @(
            @{ Name = 'QA'; CollectionID = 'MCM00101'; Purpose = 'Required'; AvailableOffsetDays = 0; DeadlineOffsetDays = 1 }))).Plan
        $rows = Expand-RingPlan -Plan $plan -Start ([datetime]'2026-10-01 08:00')
        $rows[0].CollectionID = 'MCM00999'
        $plan.Rings[0].CollectionID | Should -Be 'MCM00101'
        $plan.Rings[0].PSObject.Properties.Name | Should -Not -Contain 'AvailableDateTime'
    }
}

Describe 'Test-RingExpansion' {
    BeforeAll { $script:Now = [datetime]'2026-09-30 12:00' }

    It 'Accepts a valid expansion' {
        $rows = @(
            (New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00101' -Available ([datetime]'2026-10-01') -Deadline ([datetime]'2026-10-02')),
            (New-TestRing -Index 2 -Name 'Pilot' -CollectionID 'MCM00102' -Available ([datetime]'2026-10-03') -Deadline ([datetime]'2026-10-05'))
        )
        $v = Test-RingExpansion -Rings $rows -Now $script:Now
        $v.Errors | Should -BeNullOrEmpty
        $v.Warnings | Should -BeNullOrEmpty
    }

    It 'Refuses an edited row that shares a collection' {
        $rows = @(
            (New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00101' -Available ([datetime]'2026-10-01') -Deadline ([datetime]'2026-10-02')),
            (New-TestRing -Index 2 -Name 'Pilot' -CollectionID 'mcm00101' -Available ([datetime]'2026-10-03') -Deadline ([datetime]'2026-10-05'))
        )
        (Test-RingExpansion -Rings $rows -Now $script:Now).Errors | Should -Match 'must not share a collection'
    }

    It 'Refuses an edited row that targets a built-in collection' {
        $rows = @((New-TestRing -Index 1 -Name 'QA' -CollectionID 'SMSDM001' -Available ([datetime]'2026-10-01') -Deadline ([datetime]'2026-10-02')))
        (Test-RingExpansion -Rings $rows -Now $script:Now).Errors | Should -Match 'built-in collection SMSDM001'
    }

    It 'Refuses available after deadline, and a deadline not later than the previous ring' {
        $rows = @(
            (New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00101' -Available ([datetime]'2026-10-01') -Deadline ([datetime]'2026-10-06')),
            (New-TestRing -Index 2 -Name 'Pilot' -CollectionID 'MCM00102' -Available ([datetime]'2026-10-07') -Deadline ([datetime]'2026-10-05'))
        )
        $errors = (Test-RingExpansion -Rings $rows -Now $script:Now).Errors -join "`n"
        $errors | Should -Match "Ring 2 'Pilot' deadline 2026-10-05 00:00 must be later than its available time"
        $errors | Should -Match "must be later than the Ring 1 'QA' deadline"
    }

    It 'Refuses a deadline in the past and warns on a past available time' {
        $rows = @((New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00101' -Available ([datetime]'2026-09-29') -Deadline ([datetime]'2026-09-30 11:00')))
        $v = Test-RingExpansion -Rings $rows -Now $script:Now
        $v.Errors | Should -Match 'is in the past'
        $v.Warnings | Should -Match 'available time 2026-09-29 00:00 is in the past'
    }

    It 'Accepts dates edited as text' {
        $rows = @((New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00101' -Available '2026-10-01 08:00' -Deadline '2026-10-02 08:00'))
        (Test-RingExpansion -Rings $rows -Now $script:Now).Errors | Should -BeNullOrEmpty
    }

    It 'Skips Available rings when ordering deadlines' {
        $rows = @(
            (New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00101' -Available ([datetime]'2026-10-01') -Deadline ([datetime]'2026-10-03')),
            (New-TestRing -Index 2 -Name 'Opt-in' -CollectionID 'MCM00102' -Purpose 'Available' -Available ([datetime]'2026-10-02') -Deadline $null),
            (New-TestRing -Index 3 -Name 'Prod' -CollectionID 'MCM00103' -Available ([datetime]'2026-10-04') -Deadline ([datetime]'2026-10-06'))
        )
        (Test-RingExpansion -Rings $rows -Now $script:Now).Errors | Should -BeNullOrEmpty
    }
}

Describe 'Ring run state' {
    BeforeEach {
        $script:RunDir = Join-Path $TestDrive ('runs-' + [guid]::NewGuid().ToString('N'))
        $script:RunPath = New-TestRunFile -Folder $script:RunDir
    }

    It 'Names the run file <plan>_<objectKey>_<stamp>.json' {
        Split-Path -Path $script:RunPath -Leaf | Should -Be 'Workstation-Rings_MCM00099_20260930-100000.json'
    }

    It 'Starts every ring Held with the snapshot dates as text' {
        $run = Read-RingRun -Path $script:RunPath
        @($run.Rings).Count | Should -Be 3
        @($run.Rings | Where-Object { $_.Status -ne 'Held' }).Count | Should -Be 0
        $run.Rings[1].AvailableDateTime | Should -Be '2026-10-08T08:00:00'
        $run.Object.ID | Should -Be 'MCM00099'
        $run.RunId | Should -Be 'run-1'
        Get-RingRunNextHeld -Run $run | Should -Be 1
    }

    It 'Held -> Created records the deployment and survives a save and read' {
        $run = Read-RingRun -Path $script:RunPath
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID '{A1}' -AssignmentID 16777300 -At ([datetime]'2026-10-01 08:05'))
        Save-RingRun -Run $run -Path $script:RunPath
        $again = Read-RingRun -Path $script:RunPath
        $again.Rings[0].Status | Should -Be 'Created'
        $again.Rings[0].DeploymentID | Should -Be '{A1}'
        $again.Rings[0].AssignmentID | Should -Be '16777300'
        $again.Rings[0].CreatedAt | Should -Be '2026-10-01T08:05:00'
        Get-RingRunNextHeld -Run $again | Should -Be 2
    }

    It 'Created -> Removed is allowed' {
        $run = Read-RingRun -Path $script:RunPath
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID '{A1}')
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Removed)
        $run.Rings[0].Status | Should -Be 'Removed'
        $run.Rings[0].RemovedAt | Should -Not -BeNullOrEmpty
    }

    It 'Refuses Held -> Removed' {
        $run = Read-RingRun -Path $script:RunPath
        { Set-RingRunRingStatus -Run $run -RingIndex 2 -Status Removed } | Should -Throw '*only a Created ring can become Removed*'
    }

    It 'Refuses Created -> Created' {
        $run = Read-RingRun -Path $script:RunPath
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID '{A1}')
        { Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID '{A2}' } | Should -Throw '*only a Held ring can become Created*'
    }

    It 'Refuses Removed -> Created' {
        $run = Read-RingRun -Path $script:RunPath
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID '{A1}')
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Removed)
        { Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID '{A3}' } | Should -Throw '*only a Held ring*'
    }

    It 'Refuses Created without a deployment ID' {
        $run = Read-RingRun -Path $script:RunPath
        { Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created } | Should -Throw '*needs a deployment ID*'
    }

    It 'Refuses a ring index that is not in the run' {
        $run = Read-RingRun -Path $script:RunPath
        { Set-RingRunRingStatus -Run $run -RingIndex 9 -Status Created -DeploymentID '{A1}' } | Should -Throw '*not in this run*'
    }

    It 'Promote needs the next held ring and a Created predecessor' {
        $run = Read-RingRun -Path $script:RunPath
        (Test-RingRunPromotable -Run $run -RingIndex 1).Ok | Should -BeTrue
        (Test-RingRunPromotable -Run $run -RingIndex 2).Reason | Should -Match 'not the next held ring; ring 1 is'
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID '{A1}')
        (Test-RingRunPromotable -Run $run -RingIndex 2).Ok | Should -BeTrue
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Removed)
        (Test-RingRunPromotable -Run $run -RingIndex 2).Reason | Should -Match 'Ring 1 is Removed; Promote needs it Created'
    }

    It 'Refuses Promote from a stale file: another session promoted the ring' {
        $mine = Read-RingRun -Path $script:RunPath
        [void](Set-RingRunRingStatus -Run $mine -RingIndex 1 -Status Created -DeploymentID '{A1}')
        Save-RingRun -Run $mine -Path $script:RunPath
        $loadedByMe = Read-RingRun -Path $script:RunPath

        $other = Read-RingRun -Path $script:RunPath
        [void](Set-RingRunRingStatus -Run $other -RingIndex 2 -Status Created -DeploymentID '{B2}')
        Save-RingRun -Run $other -Path $script:RunPath

        (Test-RingRunPromotable -Run $loadedByMe -RingIndex 2).Ok | Should -BeTrue
        $fresh = Read-RingRun -Path $script:RunPath
        $check = Test-RingRunPromotable -Run $fresh -RingIndex 2
        $check.Ok | Should -BeFalse
        $check.Reason | Should -Match 'Ring 2 is Created, not Held'
    }

    It 'Refuses Promote on a closed run' {
        $run = Read-RingRun -Path $script:RunPath
        $run.ClosedAt = '2026-10-20T00:00:00'
        (Test-RingRunPromotable -Run $run -RingIndex 1).Reason | Should -Match 'closed'
    }

    It 'Is finished when no ring is Held, or when the next held ring follows a Removed ring' {
        $run = Read-RingRun -Path $script:RunPath
        Test-RingRunFinished -Run $run | Should -BeFalse
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID '{A1}')
        Test-RingRunFinished -Run $run | Should -BeFalse
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Removed)
        Test-RingRunFinished -Run $run | Should -BeTrue

        $all = Read-RingRun -Path $script:RunPath
        1..3 | ForEach-Object { [void](Set-RingRunRingStatus -Run $all -RingIndex $_ -Status Created -DeploymentID "{$_}") }
        Test-RingRunFinished -Run $all | Should -BeTrue
        Get-RingRunNextHeld -Run $all | Should -Be 0
    }

    It 'Read-RingRun refuses a file that is not a run' {
        $p = Join-Path $script:RunDir 'other.json'
        Set-Content -LiteralPath $p -Value '{"Name":"x"}' -Encoding UTF8
        { Read-RingRun -Path $p } | Should -Throw '*not a ring run*'
    }

    It 'Save-RingRun leaves no temporary file' {
        $run = Read-RingRun -Path $script:RunPath
        Save-RingRun -Run $run -Path $script:RunPath
        Test-Path -LiteralPath ($script:RunPath + '.tmp') | Should -BeFalse
    }
}

Describe 'Run-file lock' {
    It 'A second lock is refused and names the lock file; release allows a new lock' {
        $path = Join-Path $TestDrive 'lock-run.json'
        $lock = Enter-RingRunLock -Path $path
        Test-Path -LiteralPath $lock | Should -BeTrue
        { Enter-RingRunLock -Path $path } | Should -Throw "*delete $lock*"
        Exit-RingRunLock -LockPath $lock
        Test-Path -LiteralPath $lock | Should -BeFalse
        $again = Enter-RingRunLock -Path $path
        Exit-RingRunLock -LockPath $again
    }
}

Describe 'Promote after a passed deadline' {
    It 'Returns 0 when the deadline is ahead' {
        $ring = [PSCustomObject]@{ AvailableDateTime = '2026-10-08T08:00:00'; DeadlineDateTime = '2026-10-09T08:00:00' }
        Get-RingPromoteShift -Ring $ring -Now ([datetime]'2026-10-08 12:00') | Should -Be 0
    }

    It 'Shifts so the ring starts at the next minute and keeps its gap' {
        $ring = [PSCustomObject]@{ AvailableDateTime = '2026-10-08T08:00:00'; DeadlineDateTime = '2026-10-09T08:00:00' }
        $minutes = Get-RingPromoteShift -Ring $ring -Now ([datetime]'2026-10-10 09:30:20')
        $minutes | Should -Be ((([datetime]'2026-10-10 09:31') - ([datetime]'2026-10-08 08:00')).TotalMinutes)
    }

    It 'Move-RingRunSchedule shifts Held rings from the index on and records the shift' {
        $dir = Join-Path $TestDrive 'shift'
        $run = Read-RingRun -Path (New-TestRunFile -Folder $dir)
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID '{A1}')
        [void](Move-RingRunSchedule -Run $run -FromIndex 2 -Minutes 60)
        $run.Rings[0].AvailableDateTime | Should -Be '2026-10-01T08:00:00'
        $run.Rings[1].AvailableDateTime | Should -Be '2026-10-08T09:00:00'
        $run.Rings[1].DeadlineDateTime  | Should -Be '2026-10-09T09:00:00'
        $run.Rings[2].AvailableDateTime | Should -Be '2026-10-15T09:00:00'
        $run.Rings[2].ShiftedMinutes | Should -Be 60
        $run.Rings[0].ShiftedMinutes | Should -Be 0
    }
}

Describe 'Get-RingDeploymentSummary' {
    It 'Returns live counts from Get-CMDeployment -DeploymentId' {
        Mock Get-CMDeployment -ModuleName DeploymentHelperCommon {
            [PSCustomObject]@{ NumberTargeted = 50; NumberSuccess = 45; NumberErrors = 2; NumberInProgress = 3; NumberUnknown = 0; NumberOther = 0 }
        } -ParameterFilter { $DeploymentId -eq '{A1}' }
        $s = Get-RingDeploymentSummary -DeploymentId '{A1}'
        $s.Found | Should -BeTrue
        $s.Targeted | Should -Be 50
        $s.Success | Should -Be 45
        $s.Errors | Should -Be 2
        $s.InProgress | Should -Be 3
    }

    It 'Returns Found = false when the site has no such deployment' {
        Mock Get-CMDeployment -ModuleName DeploymentHelperCommon { $null }
        (Get-RingDeploymentSummary -DeploymentId '{GONE}').Found | Should -BeFalse
    }

    It 'Returns Found = null and the error when the read fails' {
        Mock Get-CMDeployment -ModuleName DeploymentHelperCommon { throw 'provider unreachable' }
        $s = Get-RingDeploymentSummary -DeploymentId '{A1}'
        $s.Found | Should -BeNullOrEmpty
        $s.Error | Should -Match 'provider unreachable'
    }

    It 'Does not call the site without a deployment ID' {
        Mock Get-CMDeployment -ModuleName DeploymentHelperCommon { throw 'should not run' }
        (Get-RingDeploymentSummary -DeploymentId '').Error | Should -Match 'No deployment ID'
        Should -Invoke Get-CMDeployment -ModuleName DeploymentHelperCommon -Times 0 -Exactly
    }
}

Describe 'Update-RingRunReconcile' {
    BeforeEach {
        $script:Run = Read-RingRun -Path (New-TestRunFile -Folder (Join-Path $TestDrive ('rec-' + [guid]::NewGuid().ToString('N'))))
        [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Created -DeploymentID '{A1}')
    }

    It 'Marks a deleted deployment Removed and reports the run finished' {
        $r = Update-RingRunReconcile -Run $script:Run -Summaries @{ 1 = [PSCustomObject]@{ Found = $false } }
        $r.Changed | Should -BeTrue
        $r.Removed | Should -Be @(1)
        $r.Finished | Should -BeTrue
        $script:Run.Rings[0].Status | Should -Be 'Removed'
    }

    It 'Leaves the status alone when the read failed' {
        $r = Update-RingRunReconcile -Run $script:Run -Summaries @{ 1 = [PSCustomObject]@{ Found = $null; Error = 'timeout' } }
        $r.Changed | Should -BeFalse
        $script:Run.Rings[0].Status | Should -Be 'Created'
        $r.Finished | Should -BeFalse
    }

    It 'Leaves a live deployment Created' {
        $r = Update-RingRunReconcile -Run $script:Run -Summaries @{ 1 = [PSCustomObject]@{ Found = $true; Targeted = 5 } }
        $r.Changed | Should -BeFalse
        $script:Run.Rings[0].Status | Should -Be 'Created'
    }
}

Describe 'Test-RingThreshold' {
    It 'Is met when no threshold is set' {
        (Test-RingThreshold -Summary $null -ThresholdPercent $null).Met | Should -BeTrue
    }

    It 'Is met at or above the threshold' {
        $s = [PSCustomObject]@{ Found = $true; Targeted = 20; Success = 18 }
        $t = Test-RingThreshold -Summary $s -ThresholdPercent 90
        $t.Met | Should -BeTrue
        $t.Percent | Should -Be 90
    }

    It 'Is not met below the threshold and says why' {
        $s = [PSCustomObject]@{ Found = $true; Targeted = 20; Success = 17 }
        $t = Test-RingThreshold -Summary $s -ThresholdPercent 90
        $t.Met | Should -BeFalse
        $t.Reason | Should -Match 'Success 85% is below the 90% threshold'
    }

    It 'Is not met with no targeted clients or no live counts' {
        (Test-RingThreshold -Summary ([PSCustomObject]@{ Found = $true; Targeted = 0; Success = 0 }) -ThresholdPercent 50).Met | Should -BeFalse
        (Test-RingThreshold -Summary ([PSCustomObject]@{ Found = $null }) -ThresholdPercent 50).Met | Should -BeFalse
    }
}

Describe 'Close-RingRun and Get-RingRunFile' {
    It 'Stamps ClosedAt and moves the file to the closed subfolder' {
        $dir = Join-Path $TestDrive 'close'
        $path = New-TestRunFile -Folder $dir
        @(Get-RingRunFile -Path $dir).Count | Should -Be 1
        $run = Read-RingRun -Path $path
        $dest = Close-RingRun -Run $run -Path $path -At ([datetime]'2026-10-20 09:00')
        Test-Path -LiteralPath $path | Should -BeFalse
        $dest | Should -Be (Join-Path (Join-Path $dir 'closed') (Split-Path $path -Leaf))
        (Read-RingRun -Path $dest).ClosedAt | Should -Be '2026-10-20T09:00:00'
        @(Get-RingRunFile -Path $dir).Count | Should -Be 0
        @(Get-RingRunFile -Path $dir -Closed).Count | Should -Be 1
    }
}

Describe 'Invoke-RingDeployment' {
    BeforeAll {
        $script:Col = [PSCustomObject]@{ Name = 'Pilot'; CollectionID = 'MCM00102'; MemberCount = 5 }
        Mock New-CMSchedule -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ Token = 'sched' } }
    }

    It 'Application: passes the ring dates and options to the existing function' {
        Mock New-CMApplicationDeployment -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ AssignmentID = 1677; AssignmentUniqueID = '{U1}' } }
        $ring = New-TestRing -Index 2 -Name 'Pilot' -CollectionID 'MCM00102' -Available '2026-10-03T08:00:00' -Deadline '2026-10-06T08:00:00' -UserNotification 'DisplaySoftwareCenterOnly'
        $ring.RebootOutsideServiceWindow = $true
        $r = Invoke-RingDeployment -Type Application -TargetObject ([PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; SoftwareVersion = '26.00' }) -Collection $script:Col -Ring $ring -TimeBasedOn Utc
        $r.Success | Should -BeTrue
        $r.DeploymentUniqueID | Should -Be '{U1}'
        Should -Invoke New-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 1 -Exactly -ParameterFilter {
            $AvailableDateTime -eq [datetime]'2026-10-03 08:00' -and $DeadlineDateTime -eq [datetime]'2026-10-06 08:00' -and
            $DeployPurpose -eq 'Required' -and $UserNotification -eq 'DisplaySoftwareCenterOnly' -and
            $TimeBaseOn -eq 'Utc' -and $RebootOutsideServiceWindow -eq $true -and $CollectionName -eq 'Pilot'
        }
    }

    It 'Package: the deadline becomes a schedule' {
        Mock New-CMPackageDeployment -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ AdvertisementID = 'MCM20001' } }
        $ring = New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00102' -Available '2026-10-01T08:00:00' -Deadline '2026-10-02T08:00:00'
        $r = Invoke-RingDeployment -Type Package -TargetObject ([PSCustomObject]@{ Name = 'Legacy'; PackageID = 'MCM00050' }) -ProgramName 'Install' -Collection $script:Col -Ring $ring
        $r.DeploymentUniqueID | Should -Be 'MCM20001'
        Should -Invoke New-CMPackageDeployment -ModuleName DeploymentHelperCommon -Times 1 -Exactly -ParameterFilter {
            $ProgramName -eq 'Install' -and $null -ne $Schedule -and -not $PSBoundParameters.ContainsKey('DeadlineDateTime')
        }
    }

    It 'SUG: maps the restart and fallback options' {
        Mock New-CMSoftwareUpdateDeployment -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ AssignmentID = 1688; AssignmentUniqueID = '{S1}' } }
        $ring = New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00102' -Available '2026-10-01T08:00:00' -Deadline '2026-10-02T08:00:00'
        $ring.RebootOutsideServiceWindow = $true
        $ring.AllowBoundaryFallback = $false
        [void](Invoke-RingDeployment -Type SUG -TargetObject ([PSCustomObject]@{ LocalizedDisplayName = '2026-10 Updates'; NumberOfUpdates = 4 }) -Collection $script:Col -Ring $ring)
        Should -Invoke New-CMSoftwareUpdateDeployment -ModuleName DeploymentHelperCommon -Times 1 -Exactly -ParameterFilter {
            $AllowRestart -eq $true -and $UnprotectedType -eq 'NoInstall' -and $DeploymentType -eq 'Required'
        }
    }

    It 'Task sequence: passes availability and progress options' {
        Mock New-CMTaskSequenceDeployment -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ AdvertisementID = 'MCM20002' } }
        $ring = New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00102' -Purpose 'Available' -Available '2026-10-01T08:00:00' -Deadline $null
        $ring.TaskSequenceAvailability = 'ClientsMediaAndPxe'
        [void](Invoke-RingDeployment -Type TaskSequence -TargetObject ([PSCustomObject]@{ Name = 'OSD'; PackageID = 'MCM00060' }) -Collection $script:Col -Ring $ring)
        Should -Invoke New-CMTaskSequenceDeployment -ModuleName DeploymentHelperCommon -Times 1 -Exactly -ParameterFilter {
            $Availability -eq 'ClientsMediaAndPxe' -and $DeployPurpose -eq 'Available' -and $null -eq $Schedule
        }
    }

    It 'Refuses a built-in collection without calling the site' {
        Mock New-CMApplicationDeployment -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ AssignmentID = 1 } }
        $ring = New-TestRing -Index 1 -Name 'QA' -CollectionID 'SMS00001' -Available '2026-10-01T08:00:00' -Deadline '2026-10-02T08:00:00'
        $col = [PSCustomObject]@{ Name = 'All Systems'; CollectionID = 'SMS00001'; MemberCount = 900 }
        $r = Invoke-RingDeployment -Type Application -TargetObject ([PSCustomObject]@{ LocalizedDisplayName = '7-Zip' }) -Collection $col -Ring $ring
        $r.Success | Should -BeFalse
        Should -Invoke New-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 0 -Exactly
    }
}

Describe 'Expansion the view submits' {
    It 'Creates one deployment per ring with the expanded dates, in ring order' {
        $script:Calls = New-Object System.Collections.Generic.List[object]
        Mock New-CMApplicationDeployment -ModuleName DeploymentHelperCommon {
            $script:Calls.Add([PSCustomObject]@{ Collection = $CollectionName; Available = $AvailableDateTime; Deadline = $DeadlineDateTime })
            [PSCustomObject]@{ AssignmentID = $script:Calls.Count; AssignmentUniqueID = ('{{U{0}}}' -f $script:Calls.Count) }
        }
        $seed = (Get-RingPlanSeed)[0]
        $plan = (ConvertTo-RingPlan -InputObject $seed).Plan
        $rows = Expand-RingPlan -Plan $plan -Start ([datetime]'2026-10-01 08:00')
        $i = 0
        foreach ($row in $rows) { $i++; $row.CollectionID = ('MCM0020{0}' -f $i) }
        (Test-RingExpansion -Rings $rows -Now ([datetime]'2026-09-30')).Errors | Should -BeNullOrEmpty

        $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; SoftwareVersion = '26.00'; PackageID = 'MCM00099' }
        $records = foreach ($row in $rows) {
            $col = [PSCustomObject]@{ Name = ('Ring ' + $row.Index); CollectionID = $row.CollectionID; MemberCount = 10 }
            $res = Invoke-RingDeployment -Type Application -TargetObject $app -Collection $col -Ring $row
            New-RingAuditRecord -Type Application -TargetObject $app -Collection $col -Ring $row -Result $res -PlanName $plan.Name -RunId 'run-9'
        }

        $script:Calls.Count | Should -Be 4
        $script:Calls[0].Available | Should -Be ([datetime]'2026-10-01 08:00')
        $script:Calls[0].Deadline  | Should -Be ([datetime]'2026-10-02 08:00')
        $script:Calls[3].Available | Should -Be ([datetime]'2026-10-15 08:00')
        $script:Calls[3].Deadline  | Should -Be ([datetime]'2026-10-18 08:00')
        $script:Calls.Collection | Should -Be @('Ring 1', 'Ring 2', 'Ring 3', 'Ring 4')

        $records[1].PlanName  | Should -Be 'Workstation-Rings'
        $records[1].RingIndex | Should -Be 2
        $records[1].RingName  | Should -Be 'Pilot'
        $records[1].RunId     | Should -Be 'run-9'
        $records[1].DeadlineDateTime | Should -Be '2026-10-06T08:00:00'
        $records[1].Result | Should -Be 'Success'
    }
}

Describe 'Test-RingPreflight' {
    BeforeAll {
        $script:Rows = @(
            (New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00101' -Available '2026-10-01T08:00:00' -Deadline '2026-10-02T08:00:00'),
            (New-TestRing -Index 2 -Name 'Pilot' -CollectionID 'MCM00102' -Available '2026-10-03T08:00:00' -Deadline '2026-10-06T08:00:00')
        )
    }

    BeforeEach {
        Mock Test-ApplicationExists -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; SoftwareVersion = '26.00'; PackageID = 'MCM00099' } }
        Mock Test-ContentDistributed -ModuleName DeploymentHelperCommon { @{ IsFullyDistributed = $true; NumberSuccess = 3; Targeted = 3 } }
        Mock Test-CollectionValid -ModuleName DeploymentHelperCommon {
            [PSCustomObject]@{ Name = ('Coll ' + $CollectionId); CollectionID = $CollectionId; CollectionType = 2; MemberCount = 10 }
        }
        Mock Test-DuplicateDeployment -ModuleName DeploymentHelperCommon { $null }
    }

    It 'Passes all five checks for every ring' {
        $p = Test-RingPreflight -Type Application -ObjectName '7-Zip' -Rings $script:Rows
        $p.Ok | Should -BeTrue
        @($p.Rings).Count | Should -Be 2
        $p.Rings[1].Message | Should -Be '5/5 checks passed'
        $p.Rings[1].Collection.CollectionID | Should -Be 'MCM00102'
        Should -Invoke Test-CollectionValid -ModuleName DeploymentHelperCommon -Times 2 -Exactly
        Should -Invoke Test-ContentDistributed -ModuleName DeploymentHelperCommon -Times 1 -Exactly
    }

    It 'Fails one ring on a duplicate and names the collection' {
        Mock Test-DuplicateDeployment -ModuleName DeploymentHelperCommon { @([PSCustomObject]@{ AssignmentID = 1 }) } -ParameterFilter { $CollectionName -eq 'Coll MCM00102' }
        $p = Test-RingPreflight -Type Application -ObjectName '7-Zip' -Rings $script:Rows
        $p.Ok | Should -BeFalse
        $p.Rings[0].Ok | Should -BeTrue
        $p.Rings[1].Message | Should -Be "4/5: A deployment of this object to 'Coll MCM00102' already exists."
    }

    It 'Fails every ring when content is not fully distributed' {
        Mock Test-ContentDistributed -ModuleName DeploymentHelperCommon { @{ IsFullyDistributed = $false; NumberSuccess = 1; Targeted = 3 } }
        $p = Test-RingPreflight -Type Application -ObjectName '7-Zip' -Rings $script:Rows
        $p.Ok | Should -BeFalse
        $p.Rings[0].Message | Should -Match 'Content is not fully distributed: 1/3'
    }

    It 'Fails a ring whose collection is not a device collection' {
        Mock Test-CollectionValid -ModuleName DeploymentHelperCommon { $null } -ParameterFilter { $CollectionId -eq 'MCM00101' }
        $p = Test-RingPreflight -Type Application -ObjectName '7-Zip' -Rings $script:Rows
        $p.Rings[0].Message | Should -Match 'MCM00101 was not found or is not a device collection'
        $p.Rings[0].Passed | Should -Be 2
    }

    It 'Refuses when the object ID no longer matches the run' {
        $p = Test-RingPreflight -Type Application -ObjectName '7-Zip' -Rings $script:Rows -ExpectedObjectId 'MCM00001'
        $p.Ok | Should -BeFalse
        $p.ObjectMessage | Should -Match 'now has ID MCM00099; the run was created for MCM00001'
    }

    It 'Refuses a package program that does not exist' {
        Mock Test-PackageExists -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ Name = 'Legacy'; PackageID = 'MCM00050' } }
        Mock Get-CMPackagePrograms -ModuleName DeploymentHelperCommon { @([PSCustomObject]@{ ProgramName = 'Install' }) }
        Mock Test-DuplicatePackageDeployment -ModuleName DeploymentHelperCommon { $null }
        (Test-RingPreflight -Type Package -ObjectName 'Legacy' -ProgramName 'Install' -Rings $script:Rows).Ok | Should -BeTrue
        $p = Test-RingPreflight -Type Package -ObjectName 'Legacy' -ProgramName 'Uninstall' -Rings $script:Rows
        $p.Ok | Should -BeFalse
        $p.ObjectMessage | Should -Match "no program 'Uninstall'"
    }

    It 'Skips the content check for a software update group' {
        Mock Test-SUGExists -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ LocalizedDisplayName = '2026-10'; CI_ID = 1001; NumberOfUpdates = 4 } }
        Mock Test-DuplicateSUGDeployment -ModuleName DeploymentHelperCommon { $null }
        (Test-RingPreflight -Type SUG -ObjectName '2026-10' -Rings $script:Rows).Ok | Should -BeTrue
        Should -Invoke Test-ContentDistributed -ModuleName DeploymentHelperCommon -Times 0 -Exactly
    }
}
Describe 'Get-RingLiveSummary' {
    BeforeAll {
        $script:Guid = '{6F9619FF-8B86-D011-B42D-00C04FC964FF}'
        $script:Ring = [PSCustomObject]@{ Index = 1; DeploymentID = $script:Guid; AssignmentID = '16777300'; CollectionName = 'Pilot' }
    }

    BeforeEach {
        Mock Get-CMDeployment -ModuleName DeploymentHelperCommon { $null } -ParameterFilter { $DeploymentId }
        Mock Get-CMDeployment -ModuleName DeploymentHelperCommon { @() } -ParameterFilter { $CollectionName }
        Mock Get-CMApplicationDeployment -ModuleName DeploymentHelperCommon { $null }
    }

    It 'Uses the summary read by deployment ID when it finds the deployment' {
        Mock Get-CMDeployment -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ NumberTargeted = 4; NumberSuccess = 4 } } -ParameterFilter { $DeploymentId -eq $script:Guid }
        $s = Get-RingLiveSummary -Ring $script:Ring -ObjectType Application
        $s.Found | Should -BeTrue
        $s.Summarized | Should -BeTrue
        $s.Success | Should -Be 4
        Should -Invoke Get-CMDeployment -ModuleName DeploymentHelperCommon -Times 0 -Exactly -ParameterFilter { $CollectionName }
        Should -Invoke Get-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 0 -Exactly
    }

    It 'Finds the summary by AssignmentID on the collection and returns the corrected ID' {
        Mock Get-CMDeployment -ModuleName DeploymentHelperCommon {
            @([PSCustomObject]@{ DeploymentID = '{REAL}'; AssignmentID = 16777300; NumberTargeted = 9; NumberSuccess = 3 })
        } -ParameterFilter { $CollectionName -eq 'Pilot' -and $FeatureType -eq 'Application' }
        $s = Get-RingLiveSummary -Ring $script:Ring -ObjectType Application
        $s.Found | Should -BeTrue
        $s.Targeted | Should -Be 9
        $s.CorrectedDeploymentID | Should -Be '{REAL}'
    }

    It 'A new deployment without a summary row stays Created: the object read finds it' {
        Mock Get-CMApplicationDeployment -ModuleName DeploymentHelperCommon {
            [PSCustomObject]@{ AssignmentUniqueID = $script:Guid }
        } -ParameterFilter { $DeploymentId -eq $script:Guid -and -not $Summary }
        $s = Get-RingLiveSummary -Ring $script:Ring -ObjectType Application
        $s.Found | Should -BeTrue
        $s.Summarized | Should -BeFalse

        $run = Read-RingRun -Path (New-TestRunFile -Folder (Join-Path $TestDrive 'lag'))
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID $script:Guid)
        $r = Update-RingRunReconcile -Run $run -Summaries @{ 1 = $s }
        $r.Changed | Should -BeFalse
        $r.Finished | Should -BeFalse
        $run.Rings[0].Status | Should -Be 'Created'
        (Test-RingThreshold -Summary $s -ThresholdPercent 90).Reason | Should -Match 'has not summarized this ring yet'
    }

    It 'Confirms a deletion only when the object read finds nothing' {
        $s = Get-RingLiveSummary -Ring $script:Ring -ObjectType Application
        $s.Found | Should -BeFalse
        Should -Invoke Get-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 1 -Exactly -ParameterFilter { $DeploymentId -eq $script:Guid }
    }

    It 'Reads the object with the per-type cmdlet for <Type>' -TestCases @(
        @{ Type = 'SUG';          Command = 'Get-CMUpdateGroupDeployment';  Id = '{6F9619FF-8B86-D011-B42D-00C04FC964FF}' }
        @{ Type = 'Package';      Command = 'Get-CMPackageDeployment';      Id = 'PS120001' }
        @{ Type = 'TaskSequence'; Command = 'Get-CMTaskSequenceDeployment'; Id = 'PS120002' }
    ) {
        Mock $Command -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ Id = 1 } }
        $ring = [PSCustomObject]@{ Index = 1; DeploymentID = $Id; AssignmentID = ''; CollectionName = 'Pilot' }
        $s = Get-RingLiveSummary -Ring $ring -ObjectType $Type
        $s.Found | Should -BeTrue
        $s.Summarized | Should -BeFalse
        Should -Invoke $Command -ModuleName DeploymentHelperCommon -Times 1 -Exactly -ParameterFilter { $DeploymentId -eq $Id }
    }

    It 'Maps SUG to the SoftwareUpdate feature type for the collection read' {
        Mock Get-CMUpdateGroupDeployment -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ Id = 1 } }
        [void](Get-RingLiveSummary -Ring $script:Ring -ObjectType SUG)
        Should -Invoke Get-CMDeployment -ModuleName DeploymentHelperCommon -Times 1 -Exactly -ParameterFilter { $FeatureType -eq 'SoftwareUpdate' }
    }

    It 'Reports uncertainty, not a deletion, when a read fails' {
        Mock Get-CMDeployment -ModuleName DeploymentHelperCommon { throw 'provider busy' } -ParameterFilter { $CollectionName }
        $s = Get-RingLiveSummary -Ring $script:Ring -ObjectType Application
        $s.Found | Should -BeNullOrEmpty
        $s.Error | Should -Match 'provider busy'

        Mock Get-CMDeployment -ModuleName DeploymentHelperCommon { @() } -ParameterFilter { $CollectionName }
        Mock Get-CMApplicationDeployment -ModuleName DeploymentHelperCommon { throw 'access denied' }
        $s = Get-RingLiveSummary -Ring $script:Ring -ObjectType Application
        $s.Found | Should -BeNullOrEmpty
        $s.Error | Should -Match 'access denied'
    }

    It 'Reports uncertainty for an ID that does not have the deployment ID form' {
        $numeric = [PSCustomObject]@{ Index = 1; DeploymentID = '16777300'; AssignmentID = ''; CollectionName = '' }
        $s = Get-RingLiveSummary -Ring $numeric -ObjectType Application
        $s.Found | Should -BeNullOrEmpty
        $s.Error | Should -Match 'cannot be confirmed'
        Should -Invoke Get-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 0 -Exactly
    }

    It 'Reconcile stores a corrected deployment ID' {
        $run = Read-RingRun -Path (New-TestRunFile -Folder (Join-Path $TestDrive 'corr'))
        [void](Set-RingRunRingStatus -Run $run -RingIndex 1 -Status Created -DeploymentID '16777300')
        $r = Update-RingRunReconcile -Run $run -Summaries @{ 1 = [PSCustomObject]@{ Found = $true; Summarized = $true; CorrectedDeploymentID = '{REAL}' } }
        $r.Changed | Should -BeTrue
        $r.Corrected | Should -Be @(1)
        $run.Rings[0].DeploymentID | Should -Be '{REAL}'
        $run.Rings[0].Status | Should -Be 'Created'
    }
}

Describe 'Get-RingNow' {
    It 'Returns UTC time for a Utc plan and local time otherwise' {
        ((Get-RingNow -TimeBasedOn Utc) - [datetime]::UtcNow).Duration().TotalSeconds | Should -BeLessThan 5
        ((Get-RingNow -TimeBasedOn LocalTime) - (Get-Date)).Duration().TotalSeconds | Should -BeLessThan 5
    }

    It 'Makes a passed UTC deadline fail expansion checks' {
        $utcNow = [datetime]::UtcNow
        $ring = New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00101' -Available $utcNow.AddHours(-2) -Deadline $utcNow.AddMinutes(-10)
        (Test-RingExpansion -Rings @($ring) -Now (Get-RingNow -TimeBasedOn Utc)).Errors | Should -Match 'is in the past'
    }
}

Describe 'Run file names never collide' {
    BeforeAll {
        $script:NewRun = {
            $rings = @((New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00101' -Available '2026-10-01T08:00:00' -Deadline '2026-10-02T08:00:00'))
            New-RingRun -PlanName 'Server-Rings' -Object ([ordered]@{ Type = 'Application'; ID = 'MCM00099'; Name = '7-Zip'; ProgramName = $null }) -Rings $rings -RunId (New-RingRunId)
        }
    }

    It 'A second run started in the same second gets its own file' {
        $dir = Join-Path $TestDrive 'collide'
        $name = 'Server-Rings_MCM00099_20260930-100000.json'
        $a = New-RingRunFile -Run (& $script:NewRun) -Folder $dir -FileName $name
        $b = New-RingRunFile -Run (& $script:NewRun) -Folder $dir -FileName $name
        $a | Should -Not -Be $b
        Split-Path $b -Leaf | Should -Be 'Server-Rings_MCM00099_20260930-100000-2.json'
        (Read-RingRun -Path $a).RunId | Should -Not -Be (Read-RingRun -Path $b).RunId
    }

    It 'A new run never takes the name of a closed run' {
        $dir = Join-Path $TestDrive 'collide-closed'
        $name = 'Server-Rings_MCM00099_20260930-100000.json'
        $a = New-RingRunFile -Run (& $script:NewRun) -Folder $dir -FileName $name
        [void](Close-RingRun -Run (Read-RingRun -Path $a) -Path $a)
        $b = New-RingRunFile -Run (& $script:NewRun) -Folder $dir -FileName $name
        Split-Path $b -Leaf | Should -Be 'Server-Rings_MCM00099_20260930-100000-2.json'
    }

    It 'Closing never replaces an existing closed file' {
        $dir = Join-Path $TestDrive 'collide-close'
        $name = 'Server-Rings_MCM00099_20260930-100000.json'
        $first = New-RingRunFile -Run (& $script:NewRun) -Folder $dir -FileName $name
        $firstId = (Read-RingRun -Path $first).RunId
        $closedFirst = Close-RingRun -Run (Read-RingRun -Path $first) -Path $first
        $copy = Join-Path $dir $name
        Copy-Item -LiteralPath $closedFirst -Destination $copy
        $closedSecond = Close-RingRun -Run (Read-RingRun -Path $copy) -Path $copy
        $closedSecond | Should -Not -Be $closedFirst
        (Read-RingRun -Path $closedFirst).RunId | Should -Be $firstId
        @(Get-RingRunFile -Path $dir -Closed).Count | Should -Be 2
    }
}

Describe 'Creating state' {
    BeforeEach {
        $script:RunPath = New-TestRunFile -Folder (Join-Path $TestDrive ('creating-' + [guid]::NewGuid().ToString('N')))
        $script:Run = Read-RingRun -Path $script:RunPath
    }

    It 'Held -> Creating records CreatingAt and clears the IDs' {
        [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Creating -At ([datetime]'2026-10-01 08:00'))
        $script:Run.Rings[0].Status | Should -Be 'Creating'
        $script:Run.Rings[0].CreatingAt | Should -Be '2026-10-01T08:00:00'
        $script:Run.Rings[0].DeploymentID | Should -BeNullOrEmpty
    }

    It 'Creating -> Created records the deployment and clears CreatingAt' {
        [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Creating)
        [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Created -DeploymentID '{A1}' -AssignmentID 16777300)
        $script:Run.Rings[0].Status | Should -Be 'Created'
        $script:Run.Rings[0].DeploymentID | Should -Be '{A1}'
        $script:Run.Rings[0].CreatingAt | Should -BeNullOrEmpty
    }

    It 'Creating -> Held clears CreatingAt' {
        [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Creating)
        [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Held)
        $script:Run.Rings[0].Status | Should -Be 'Held'
        $script:Run.Rings[0].CreatingAt | Should -BeNullOrEmpty
    }

    It 'Refuses <From> -> <To>' -TestCases @(
        @{ From = 'Created';  To = 'Creating'; Message = 'only a Held ring can become Creating' }
        @{ From = 'Held';     To = 'Held';     Message = 'only a Creating ring can return to Held' }
        @{ From = 'Removed';  To = 'Creating'; Message = 'only a Held ring can become Creating' }
        @{ From = 'Creating'; To = 'Removed';  Message = 'only a Created ring can become Removed' }
    ) {
        switch ($From) {
            'Created'  { [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Created -DeploymentID '{A1}') }
            'Removed'  { [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Created -DeploymentID '{A1}'); [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Removed) }
            'Creating' { [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Creating) }
        }
        { Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status $To -DeploymentID '{A2}' } | Should -Throw "*$Message*"
    }

    It 'A Creating ring is the next pending ring, blocks Promote, and keeps the run open' {
        [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Created -DeploymentID '{A1}')
        [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 2 -Status Creating)
        Get-RingRunNextHeld -Run $script:Run | Should -Be 2
        (Test-RingRunPromotable -Run $script:Run -RingIndex 2).Reason | Should -Match 'unresolved create attempt'
        (Test-RingRunPromotable -Run $script:Run -RingIndex 3).Reason | Should -Match 'ring 2 is'
        Test-RingRunFinished -Run $script:Run | Should -BeFalse
    }

    It 'A stale file is refused: another session started creating the ring' {
        $mine = Read-RingRun -Path $script:RunPath
        [void](Set-RingRunRingStatus -Run $mine -RingIndex 1 -Status Created -DeploymentID '{A1}')
        Save-RingRun -Run $mine -Path $script:RunPath
        $other = Read-RingRun -Path $script:RunPath
        [void](Set-RingRunRingStatus -Run $other -RingIndex 2 -Status Creating)
        Save-RingRun -Run $other -Path $script:RunPath
        (Test-RingRunPromotable -Run (Read-RingRun -Path $script:RunPath) -RingIndex 2).Ok | Should -BeFalse
    }

    It 'CreatingAt survives a save and read' {
        [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Creating -At ([datetime]'2026-10-01 08:00'))
        Save-RingRun -Run $script:Run -Path $script:RunPath
        (Read-RingRun -Path $script:RunPath).Rings[0].CreatingAt | Should -Be '2026-10-01T08:00:00'
    }
}

Describe 'Update-RingRunReconcile resolves pending creates' {
    BeforeEach {
        $script:Run = Read-RingRun -Path (New-TestRunFile -Folder (Join-Path $TestDrive ('pending-' + [guid]::NewGuid().ToString('N'))))
        [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 1 -Status Created -DeploymentID '{A1}')
        [void](Set-RingRunRingStatus -Run $script:Run -RingIndex 2 -Status Creating -At ([datetime]'2026-10-08 08:00'))
        $script:Live = @{ 1 = [PSCustomObject]@{ Found = $true; Summarized = $true } }
    }

    It 'Adopts the one matching deployment' {
        $candidates = @{ 2 = @{ CheckFailed = $false; Error = ''; Deployments = @([PSCustomObject]@{ AssignmentUniqueID = '{NEW}'; AssignmentID = 16777500 }) } }
        $r = Update-RingRunReconcile -Run $script:Run -Summaries $script:Live -CreatingCandidates $candidates -At ([datetime]'2026-10-08 08:01')
        $r.Recovered | Should -Be @(2)
        $r.Changed | Should -BeTrue
        $script:Run.Rings[1].Status | Should -Be 'Created'
        $script:Run.Rings[1].DeploymentID | Should -Be '{NEW}'
        $script:Run.Rings[1].AssignmentID | Should -Be '16777500'
    }

    It 'Keeps the ring pending when two deployments match' {
        $two = @([PSCustomObject]@{ AssignmentUniqueID = '{X}' }, [PSCustomObject]@{ AssignmentUniqueID = '{Y}' })
        $r = Update-RingRunReconcile -Run $script:Run -Summaries $script:Live -CreatingCandidates @{ 2 = @{ CheckFailed = $false; Deployments = $two } } -At ([datetime]'2026-10-08 09:00')
        $r.Changed | Should -BeFalse
        $script:Run.Rings[1].Status | Should -Be 'Creating'
        $r.RecoveryNotes[2] | Should -Match 'Found 2 matching deployments'
    }

    It 'Keeps the ring pending with no match inside the wait' {
        $r = Update-RingRunReconcile -Run $script:Run -Summaries $script:Live -CreatingCandidates @{ 2 = @{ CheckFailed = $false; Deployments = @() } } -At ([datetime]'2026-10-08 08:04')
        $r.Changed | Should -BeFalse
        $script:Run.Rings[1].Status | Should -Be 'Creating'
    }

    It 'Returns the ring to Held with no match after the wait' {
        $r = Update-RingRunReconcile -Run $script:Run -Summaries $script:Live -CreatingCandidates @{ 2 = @{ CheckFailed = $false; Deployments = @() } } -At ([datetime]'2026-10-08 08:05')
        $r.Reset | Should -Be @(2)
        $r.Changed | Should -BeTrue
        $script:Run.Rings[1].Status | Should -Be 'Held'
        $r.Finished | Should -BeFalse
    }

    It 'Changes nothing when the check failed or did not run' {
        $r = Update-RingRunReconcile -Run $script:Run -Summaries $script:Live -CreatingCandidates @{ 2 = @{ CheckFailed = $true; Error = 'provider busy'; Deployments = @() } } -At ([datetime]'2026-10-09')
        $r.Changed | Should -BeFalse
        $r.RecoveryNotes[2] | Should -Match 'provider busy'
        $r = Update-RingRunReconcile -Run $script:Run -Summaries $script:Live -At ([datetime]'2026-10-09')
        $r.Changed | Should -BeFalse
        $script:Run.Rings[1].Status | Should -Be 'Creating'
    }

    It 'Keeps the ring pending when the match has no deployment ID' {
        $r = Update-RingRunReconcile -Run $script:Run -Summaries $script:Live -CreatingCandidates @{ 2 = @{ CheckFailed = $false; Deployments = @([PSCustomObject]@{ Name = 'x' }) } } -At ([datetime]'2026-10-08 08:01')
        $r.Changed | Should -BeFalse
        $r.RecoveryNotes[2] | Should -Match 'Could not determine the deployment ID'
    }
}

Describe 'Get-RingDeploymentReference' {
    It 'Reads <Property> for <Type>' -TestCases @(
        @{ Type = 'Application';  Property = 'AssignmentUniqueID' }
        @{ Type = 'SUG';          Property = 'AssignmentUniqueID' }
        @{ Type = 'Package';      Property = 'AdvertisementID' }
        @{ Type = 'TaskSequence'; Property = 'AdvertisementID' }
    ) {
        $d = [PSCustomObject]@{ $Property = 'ID-1'; AssignmentID = 7 }
        $ref = & (Get-Module DeploymentHelperCommon) { param($t, $o) Get-RingDeploymentReference -Type $t -Deployment $o } $Type $d
        $ref.DeploymentID | Should -Be 'ID-1'
        $ref.AssignmentID | Should -Be '7'
    }

    It 'Throws when the object has no deployment ID' {
        { & (Get-Module DeploymentHelperCommon) { Get-RingDeploymentReference -Type Application -Deployment ([PSCustomObject]@{ Name = 'x' }) } } | Should -Throw '*Could not determine the deployment ID*'
    }
}

Describe 'Test-RingPreflight duplicate details and content skip' {
    BeforeAll {
        $script:Rows2 = @((New-TestRing -Index 1 -Name 'QA' -CollectionID 'MCM00101' -Available '2026-10-01T08:00:00' -Deadline '2026-10-02T08:00:00'))
    }

    BeforeEach {
        Mock Test-ApplicationExists -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; SoftwareVersion = '26.00'; PackageID = 'MCM00099' } }
        Mock Test-ContentDistributed -ModuleName DeploymentHelperCommon { @{ IsFullyDistributed = $false; NumberSuccess = 2; Targeted = 3 } }
        Mock Test-CollectionValid -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ Name = 'Coll'; CollectionID = $CollectionId; CollectionType = 2; MemberCount = 10 } }
        Mock Test-DuplicateDeployment -ModuleName DeploymentHelperCommon { $null }
    }

    It '-SkipContentCheck passes content in progress without reading it' {
        $p = Test-RingPreflight -Type Application -ObjectName '7-Zip' -Rings $script:Rows2 -SkipContentCheck
        $p.Ok | Should -BeTrue
        Should -Invoke Test-ContentDistributed -ModuleName DeploymentHelperCommon -Times 0 -Exactly
        (Test-RingPreflight -Type Application -ObjectName '7-Zip' -Rings $script:Rows2).Ok | Should -BeFalse
    }

    It 'A failed duplicate check fails the ring and names the error' {
        Mock Test-DuplicateDeployment -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ DuplicateCheckFailed = $true; Error = 'SMS Provider timed out' } }
        $p = Test-RingPreflight -Type Application -ObjectName '7-Zip' -Rings $script:Rows2 -SkipContentCheck
        $p.Ok | Should -BeFalse
        $p.Rings[0].DuplicateCheckError | Should -Be 'SMS Provider timed out'
        $p.Rings[0].Message | Should -Match 'Could not check for an existing deployment: SMS Provider timed out'
        $p.Rings[0].ExistingDeployments | Should -BeNullOrEmpty
    }

    It 'An existing deployment is returned for recovery' {
        Mock Test-DuplicateDeployment -ModuleName DeploymentHelperCommon { @([PSCustomObject]@{ AssignmentUniqueID = '{E1}' }) }
        $p = Test-RingPreflight -Type Application -ObjectName '7-Zip' -Rings $script:Rows2 -SkipContentCheck
        $p.Rings[0].CollectionSafe | Should -BeTrue
        @($p.Rings[0].ExistingDeployments).Count | Should -Be 1
        $p.Rings[0].ExistingDeployments[0].AssignmentUniqueID | Should -Be '{E1}'
    }
}
