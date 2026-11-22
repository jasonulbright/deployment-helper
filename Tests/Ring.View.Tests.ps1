#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'The cut-out shell functions read these stand-in controls through scope.')]
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'The shell keeps its preferences in globals; the tests set the same names.')]
param ()
# The ring preview grid and the Open Runs grid reach the screen. The functions of the script that
# fill them run here with a plan that the module loads from a file and run files that the module
# writes, with real WPF controls (not shown). The test reads the rows that each grid holds and
# the cell values that the user sees. A test that checks only that a control exists, or that a
# function was called, passes while a grid is empty.
#
# The script cannot be dot-sourced (it opens the window), so each function is cut out of its
# syntax tree and defined here.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase -ErrorAction SilentlyContinue
    # WPF controls need an STA thread. Windows PowerShell and PowerShell 7 start in STA; a harness that starts MTA must add -STA.
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { throw 'These tests create WPF controls and need an STA thread. Start the host with -STA.' }
    Import-AppModule
    Initialize-Logging -LogPath (Join-Path $TestDrive 'ring-view.log')
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:AppRoot 'start-deploymenthelper.ps1'), [ref]$null, [ref]$null)
    foreach ($n in 'Update-RingPreview', 'Update-RingPreviewGrid', 'Update-RingRunsGrid', 'Get-RingSelectedPlan', 'Reset-RingValidation', 'Format-RingDate') {
        $def = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true)
        if (-not $def) { throw "function $n not found in the script" }
        . ([scriptblock]::Create($def.Extent.Text.Replace('$PSScriptRoot', '$script:AppRoot')))
    }

    # The plan that Update Planner writes, loaded the way the Ring view loads its plan folder.
    $script:PlanDir = Join-Path $TestDrive 'plans'
    New-Item -ItemType Directory -Path $script:PlanDir -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Fixtures\shared-ring-plan.json') -Destination $script:PlanDir
    $script:Loaded = Get-RingPlanList -Path $script:PlanDir
}

AfterAll {
    Remove-Variable -Name Prefs -Scope Global -ErrorAction SilentlyContinue
}

Describe 'The ring preview grid shows the rings of the selected plan' {
    BeforeEach {
        $script:RingPlans = @($script:Loaded.Plans)
        $script:RingValidated = 'something'
        $script:RingPreviewRows = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
        $dgRingPreview = New-Object System.Windows.Controls.DataGrid
        $dgRingPreview.CanUserAddRows = $false
        $dgRingPreview.ItemsSource = $script:RingPreviewRows
        $btnRingCreate = [PSCustomObject]@{ Visibility = [System.Windows.Visibility]::Visible }
        $cboRingPlan = [PSCustomObject]@{ SelectedIndex = 0 }
        $dtpRingStart = [PSCustomObject]@{ SelectedDateTime = [datetime]'2026-12-01T08:00:00' }
    }

    It 'loads the plan file' {
        @($script:Loaded.Plans).Count | Should -Be 1
        @($script:Loaded.Errors).Count | Should -Be 0
    }

    It 'shows one row for each ring of the plan, with the name, the collection and the dates as text' {
        $fixture = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Fixtures\shared-ring-plan.json') -Raw | ConvertFrom-Json
        Update-RingPreview
        $dgRingPreview.Items.Count | Should -Be @($fixture.Rings).Count
        $dgRingPreview.Items.Count | Should -BeGreaterThan 1
        $dgRingPreview.Items[0].Name | Should -Be $fixture.Rings[0].Name
        $dgRingPreview.Items[0].CollectionID | Should -Be $fixture.Rings[0].CollectionID
        $dgRingPreview.Items[0].AvailableText | Should -Be '2026-12-01 08:00'
        $dgRingPreview.Items[0].DeadlineText | Should -Match '^2026-12-04 '
        @($dgRingPreview.Items | Where-Object { [string]::IsNullOrWhiteSpace($_.AvailableText) }).Count | Should -Be 0
    }

    It 'hides the Create button and clears the validated plan when the rows change' {
        Update-RingPreview
        $script:RingValidated | Should -BeNullOrEmpty
        [string]$btnRingCreate.Visibility | Should -Be 'Collapsed'
    }

    It 'shows no row when no plan is selected' {
        Update-RingPreview
        $dgRingPreview.Items.Count | Should -BeGreaterThan 0
        $cboRingPlan.SelectedIndex = -1
        Update-RingPreview
        $dgRingPreview.Items.Count | Should -Be 0
    }
}

Describe 'The Open Runs grid shows the rings of the run files' {
    BeforeEach {
        $script:RingRunStatePath = Join-Path $TestDrive ('runs-{0}' -f [guid]::NewGuid().ToString('N'))
        $script:RingPlans = @($script:Loaded.Plans)
        $script:RingPreviewRows = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
        $dgRingRuns = New-Object System.Windows.Controls.DataGrid
        $dgRingRuns.CanUserAddRows = $false
        $lblRingRuns = New-Object System.Windows.Controls.TextBlock
        $chkRingShowClosed = [PSCustomObject]@{ IsChecked = $false }
        $rings = @(foreach ($r in (Expand-RingPlan -Plan $script:Loaded.Plans[0] -Start ([datetime]'2026-12-01T08:00:00'))) { $r })
        $script:RingCount = $rings.Count
        $object = [ordered]@{ Type = 'Application'; ID = 'MCM00099'; Name = '7-Zip'; ProgramName = $null }
        $run = New-RingRun -PlanName 'UpdatePlanner-a1b2c3d4' -Object $object -Rings $rings -RunId 'run-view-1' -CreatedAt ([datetime]'2026-11-30T10:00:00')
        $run.Rings[0].Status = 'Created'
        $run.Rings[0].DeploymentID = 'MCM20001'
        $script:RunPath = New-RingRunFile -Run $run -Folder $script:RingRunStatePath -FileName 'UpdatePlanner-a1b2c3d4_MCM00099_20261130-100000.json'
    }

    It 'shows one row for each ring of the open run, with the status and the deadline text' {
        Update-RingRunsGrid
        $dgRingRuns.Items.Count | Should -Be $script:RingCount
        $dgRingRuns.Items[0].Status | Should -Be 'Created'
        $dgRingRuns.Items[1].Status | Should -Be 'Held'
        $dgRingRuns.Items[0].Name | Should -Not -BeNullOrEmpty
        $dgRingRuns.Items[0].CollectionID | Should -Match '^MCM'
        $dgRingRuns.Items[0].DeadlineText | Should -Match '^2026-12-04 '
        $dgRingRuns.Items[0].RunLabel | Should -Match '7-Zip'
        $lblRingRuns.Text | Should -Be 'Open Runs'
    }

    It 'tells the user to refresh for live counts on a created ring, and offers Promote on the next held ring' {
        Update-RingRunsGrid
        $dgRingRuns.Items[0].Note | Should -Be 'Refresh to read live counts.'
        $dgRingRuns.Items[1].ShowPromote | Should -BeTrue
        $dgRingRuns.Items[1].CanPromote | Should -BeTrue
        $dgRingRuns.Items[2].ShowPromote | Should -BeFalse
    }

    It 'shows a row with the reason for a run file that cannot be read, not a blank grid' {
        Set-Content -LiteralPath (Join-Path $script:RingRunStatePath 'broken_run_20261130-110000.json') -Value '{ not json'
        Update-RingRunsGrid
        $dgRingRuns.Items.Count | Should -Be ($script:RingCount + 1)
        $broken = @($dgRingRuns.Items | Where-Object { $_.RunLabel -like 'broken_run*' })
        $broken.Count | Should -Be 1
        $broken[0].Note | Should -Not -BeNullOrEmpty
    }

    It 'shows an empty grid when there is no run file' {
        Remove-Item -LiteralPath $script:RunPath -Force
        Update-RingRunsGrid
        $dgRingRuns.Items.Count | Should -Be 0
    }

    It 'shows the rings of the closed runs, and not those of the open runs, when the closed box is checked' {
        $closed = Join-Path $script:RingRunStatePath 'closed'
        $rings = @(foreach ($r in (Expand-RingPlan -Plan $script:Loaded.Plans[0] -Start ([datetime]'2026-12-01T08:00:00'))) { $r })
        $run = New-RingRun -PlanName 'Closed-Plan' -Object ([ordered]@{ Type = 'Application'; ID = 'MCM00077'; Name = 'Old App'; ProgramName = $null }) -Rings $rings -RunId 'run-closed-1' -CreatedAt ([datetime]'2026-10-01T10:00:00')
        $null = New-RingRunFile -Run $run -Folder $closed -FileName 'Closed-Plan_MCM00077_20261001-100000.json'
        $chkRingShowClosed.IsChecked = $true
        Update-RingRunsGrid
        $lblRingRuns.Text | Should -Be 'Closed Runs'
        $dgRingRuns.Items.Count | Should -Be $script:RingCount
        @($dgRingRuns.Items | Where-Object { $_.RunLabel -match 'Old App' }).Count | Should -Be $script:RingCount
        @($dgRingRuns.Items | Where-Object { $_.RunLabel -match '7-Zip' }).Count | Should -Be 0
        @($dgRingRuns.Items | Where-Object { $_.ShowPromote }).Count | Should -Be 0
    }

    It 'offers Promote on ring 1 of a run where every ring is held' {
        $rings = @(foreach ($r in (Expand-RingPlan -Plan $script:Loaded.Plans[0] -Start ([datetime]'2026-12-01T08:00:00'))) { $r })
        $run = New-RingRun -PlanName 'Held-Plan' -Object ([ordered]@{ Type = 'Application'; ID = 'MCM00055'; Name = 'Held App'; ProgramName = $null }) -Rings $rings -RunId 'run-held-1' -CreatedAt ([datetime]'2026-11-30T11:00:00')
        $null = New-RingRunFile -Run $run -Folder $script:RingRunStatePath -FileName 'Held-Plan_MCM00055_20261130-110000.json'
        Update-RingRunsGrid
        $held = @($dgRingRuns.Items | Where-Object { $_.RunLabel -match 'Held App' })
        $held.Count | Should -Be $script:RingCount
        $held[0].Status | Should -Be 'Held'
        $held[0].ShowPromote | Should -BeTrue
        $held[0].CanPromote | Should -BeTrue
        $held[0].PromoteTip | Should -Be 'Create ring 1.'
        $held[1].ShowPromote | Should -BeFalse
    }
}

Describe 'The ring grids bind to properties that the rows have' {
    BeforeAll {
        [xml]$script:Xaml = Get-Content -LiteralPath (Join-Path $script:AppRoot 'MainWindow.xaml') -Raw
        $script:XamlNs = New-Object System.Xml.XmlNamespaceManager($script:Xaml.NameTable)
        $script:XamlNs.AddNamespace('w', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
        $script:XamlNs.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
        function Get-GridBindingPath {
            param([string]$GridName)
            $node = $script:Xaml.SelectSingleNode(('//w:DataGrid[@x:Name=''{0}'']' -f $GridName), $script:XamlNs)
            if (-not $node) { throw "grid $GridName not found in MainWindow.xaml" }
            @([regex]::Matches($node.OuterXml, '\{Binding\s+(?:Path=)?([A-Za-z_][A-Za-z0-9_]*)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        }
    }

    It 'gives the preview grid a property for every column binding' {
        $script:RingPlans = @($script:Loaded.Plans)
        $script:RingValidated = $null
        $script:RingPreviewRows = New-Object System.Collections.ObjectModel.ObservableCollection[PSObject]
        $btnRingCreate = [PSCustomObject]@{ Visibility = [System.Windows.Visibility]::Visible }
        $cboRingPlan = [PSCustomObject]@{ SelectedIndex = 0 }
        $dtpRingStart = [PSCustomObject]@{ SelectedDateTime = [datetime]'2026-12-01T08:00:00' }
        $dgRingPreview = New-Object System.Windows.Controls.DataGrid
        $dgRingPreview.ItemsSource = $script:RingPreviewRows
        Update-RingPreview
        $paths = @(Get-GridBindingPath -GridName 'dgRingPreview')
        $paths.Count | Should -BeGreaterThan 3
        foreach ($path in $paths) { $null -ne $dgRingPreview.Items[0].PSObject.Properties[$path] | Should -BeTrue -Because "the column binds to $path" }
    }

    It 'gives the Open Runs grid a property for every column binding' {
        $script:RingRunStatePath = Join-Path $TestDrive ('runs-{0}' -f [guid]::NewGuid().ToString('N'))
        $rings = @(foreach ($r in (Expand-RingPlan -Plan $script:Loaded.Plans[0] -Start ([datetime]'2026-12-01T08:00:00'))) { $r })
        $run = New-RingRun -PlanName 'Bind-Plan' -Object ([ordered]@{ Type = 'Application'; ID = 'MCM00099'; Name = '7-Zip'; ProgramName = $null }) -Rings $rings -RunId 'run-bind-1' -CreatedAt ([datetime]'2026-11-30T10:00:00')
        $null = New-RingRunFile -Run $run -Folder $script:RingRunStatePath -FileName 'Bind-Plan_MCM00099_20261130-100000.json'
        $dgRingRuns = New-Object System.Windows.Controls.DataGrid
        $lblRingRuns = New-Object System.Windows.Controls.TextBlock
        $chkRingShowClosed = [PSCustomObject]@{ IsChecked = $false }
        Update-RingRunsGrid
        $paths = @(Get-GridBindingPath -GridName 'dgRingRuns')
        $paths.Count | Should -BeGreaterThan 3
        foreach ($path in $paths) { $null -ne $dgRingRuns.Items[0].PSObject.Properties[$path] | Should -BeTrue -Because "the column binds to $path" }
    }
}
