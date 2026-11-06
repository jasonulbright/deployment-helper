#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }

<#
.SYNOPSIS
    Contract test for the Suite ring plan that Update Planner writes.

.DESCRIPTION
    Loads the shared fixture (the same file is in the Update Planner repository)
    with Import-RingPlan, as a ring plan file. The extra fields (Format,
    SchemaVersion, UpdatePlanner and the Updates list of each ring) must not
    stop the load. Deployment Helper needs no code change for them.
    Run: Invoke-Pester -Path .\Tests\RingPlanContract.Tests.ps1 -Output Detailed
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Module\DeploymentHelperCommon.psd1') -Force
    Initialize-Logging -LogPath (Join-Path $TestDrive 'contract.log')
    $script:FixturePath = Join-Path $PSScriptRoot 'Fixtures\shared-ring-plan.json'
    $script:SchemaPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'docs\schema\ring-plan.schema.json'

    function Get-NormalizedHash {
        param([string]$Path)
        $text = [IO.File]::ReadAllText($Path).Replace("`r`n", "`n")
        [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($text))).Replace('-', '').ToLowerInvariant()
    }
}

Describe 'Suite ring plan contract' {
    It 'keeps the same fixture and schema as Update Planner' {
        # The hashes pin the files of Update Planner v2026.11.06.0003 (docs/schema and Tests/Fixtures).
        Get-NormalizedHash -Path $script:FixturePath | Should -Be 'd2ea422e6b647ed1c90e619ba6ddd7330ae9826e2e87d5dab1e3e16bf1f85964'
        Get-NormalizedHash -Path $script:SchemaPath | Should -Be '15546c44bb86c9da095d2873986b93ceedc45ea3c62f3cc266a24bec43a3dba8'
    }

    It 'loads the fixture as a ring plan with Import-RingPlan' {
        $plan = Import-RingPlan -Path $script:FixturePath
        $plan.Name | Should -Be 'UpdatePlanner-a1b2c3d4'
        $plan.TimeBasedOn | Should -Be 'LocalTime'
        $plan.HoldLaterRings | Should -BeFalse
        @($plan.Rings).Count | Should -Be 3
        $plan.Rings[0].Name | Should -Be 'Wave 1 - Zone 01'
        $plan.Rings[0].CollectionID | Should -Be 'MCM00346'
        $plan.Rings[0].CollectionName | Should -Be 'Zone 01'
        $plan.Rings[0].Purpose | Should -Be 'Required'
        @($plan.Rings | ForEach-Object { $_.AvailableOffsetDays }) | Should -Be @(0, 7, 14)
        @($plan.Rings | ForEach-Object { $_.DeadlineOffsetDays }) | Should -Be @(3, 10, 17)
        $plan.Rings[0].UserNotification | Should -Be 'DisplayAll'
        $plan.Rings[0].OverrideServiceWindow | Should -BeFalse
        $plan.Rings[0].SuccessThresholdPercent | Should -BeNullOrEmpty
    }

    It 'ignores the extra fields and does not carry them into the plan' {
        $plan = Import-RingPlan -Path $script:FixturePath
        $plan.PSObject.Properties.Name | Should -Not -Contain 'Format'
        $plan.PSObject.Properties.Name | Should -Not -Contain 'UpdatePlanner'
        $plan.Rings[0].PSObject.Properties.Name | Should -Not -Contain 'Updates'
    }

    It 'expands the fixture to dated deployments' {
        $plan = Import-RingPlan -Path $script:FixturePath
        $expanded = Expand-RingPlan -Plan $plan -Start ([datetime]'2026-12-01')
        $expanded.Count | Should -Be 3
        $expanded[2].CollectionID | Should -Be 'MCM00348'
        $expanded[0].DeadlineDateTime | Should -Be ([datetime]'2026-12-04')
        $expanded[2].AvailableDateTime | Should -Be ([datetime]'2026-12-15')
    }

    It 'validates the fixture against the schema' -Skip:($PSVersionTable.PSVersion.Major -lt 7) {
        $strSchema = Get-Content -LiteralPath $script:SchemaPath -Raw
        { $strSchema | ConvertFrom-Json -ErrorAction Stop } | Should -Not -Throw -Because 'a malformed schema must fail, not skip'
        try { '{}' | Test-Json -Schema $strSchema -ErrorAction Stop | Out-Null } catch { if ($_.Exception.Message -like '*parse the JSON schema*') { Set-ItResult -Skipped -Because 'This PowerShell version cannot parse the schema.'; return } }
        (Get-Content -LiteralPath $script:FixturePath -Raw) | Test-Json -Schema $strSchema | Should -BeTrue
    }
}
