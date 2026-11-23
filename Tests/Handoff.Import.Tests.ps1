#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }

<#
.SYNOPSIS
    Tests of the two imports that show their source and age (H-1, G-7): the plan from Update Planner and the result of Check ConfigMgr.

.DESCRIPTION
    Covers the handoff reader (header, age, stale limit, site, malformed and newer files), the load of the plan as a ring plan,
    the read-only list of planned updates, the copy into the plan list, and the source list of the Check ConfigMgr result.
    Site codes are fictional (AB1, CD2). Nothing touches the real profile.
    Run: Invoke-Pester -Path .\Tests\Handoff.Import.Tests.ps1 -Output Detailed
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\Module\DeploymentHelperCommon.psd1') -Force
    Initialize-Logging -LogPath (Join-Path $TestDrive 'handoff.log')
    $script:Fixture = Join-Path $PSScriptRoot 'Fixtures\shared-ring-plan.json'
    $script:Now = [datetime]::UtcNow

    function New-HandoffFile {
        param([string]$Path, [datetime]$Created = ([datetime]::UtcNow.AddHours(-2)), [string]$Site = 'AB1', [switch]$NoHeader, [int]$Version = 1, [string]$Producer = 'UpdatePlanner')
        $doc = Get-Content -LiteralPath $script:Fixture -Raw | ConvertFrom-Json
        $doc.SchemaVersion = $Version
        if (-not $NoHeader) {
            $doc | Add-Member -NotePropertyName Producer -NotePropertyValue ([pscustomobject]@{ Name = $Producer; Version = '2026.11.23.0017' }) -Force
            $doc | Add-Member -NotePropertyName CreatedUtc -NotePropertyValue ($Created.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")) -Force
            $doc | Add-Member -NotePropertyName Site -NotePropertyValue ([pscustomobject]@{ SiteCode = $Site }) -Force
        }
        $dir = Split-Path $Path -Parent
        if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
        [IO.File]::WriteAllText($Path, ($doc | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))
    }
}

Describe 'Stale limit from the prefs' {
    It 'defaults to 3 days' {
        ConvertTo-DhStaleHours -Prefs @{} | Should -Be 72
        ConvertTo-DhStaleHours -Prefs $null | Should -Be 72
    }
    It 'reads hours and days' {
        ConvertTo-DhStaleHours -Prefs @{ HandoffStaleValue = 12; HandoffStaleUnit = 'hours' } | Should -Be 12
        ConvertTo-DhStaleHours -Prefs @{ HandoffStaleValue = 2; HandoffStaleUnit = 'days' } | Should -Be 48
        ConvertTo-DhStaleHours -Prefs ([pscustomobject]@{ HandoffStaleValue = 5; HandoffStaleUnit = 'days' }) | Should -Be 120
    }
    It 'falls back to the default for a bad value' {
        ConvertTo-DhStaleHours -Prefs @{ HandoffStaleValue = 0; HandoffStaleUnit = 'days' } | Should -Be 72
        ConvertTo-DhStaleHours -Prefs @{ HandoffStaleValue = 'x'; HandoffStaleUnit = 'days' } | Should -Be 72
        ConvertTo-DhStaleHours -Prefs @{ HandoffStaleValue = 4; HandoffStaleUnit = 'weeks' } | Should -Be 72
    }
}

Describe 'Get-DhHandoffPlan' {
    BeforeEach { $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N')) }

    It 'reports a missing file as Missing and writes nothing' {
        $r = Get-DhHandoffPlan -SiteFolder 'AB1' -DataRoot $script:Root
        $r.Status | Should -Be 'Missing'
        Test-Path -LiteralPath $script:Root | Should -BeFalse
    }
    It 'reads the plan from the handoff folder of the site with the source, the site and the age' {
        New-HandoffFile -Path (Get-SuiteHandoffPath -SiteCode 'AB1' -Kind 'ring-plan' -DataRoot $script:Root)
        $r = Get-DhHandoffPlan -SiteFolder 'AB1' -DataRoot $script:Root -StaleAfterHours 72
        $r.Status | Should -Be 'Ok'
        $r.ProducerName | Should -Be 'UpdatePlanner'
        $r.ProducerVersion | Should -Be '2026.11.23.0017'
        $r.SiteCode | Should -Be 'AB1'
        $r.AgeText | Should -Be '2 hours old'
        $r.IsStale | Should -BeFalse
        $r.PlanName | Should -Be 'UpdatePlanner-a1b2c3d4'
        $r.RingCount | Should -Be 3
        @($r.Updates).Count | Should -BeGreaterThan 0
    }
    It 'does not read the folder of another site' {
        New-HandoffFile -Path (Get-SuiteHandoffPath -SiteCode 'AB1' -Kind 'ring-plan' -DataRoot $script:Root)
        (Get-DhHandoffPlan -SiteFolder 'CD2' -DataRoot $script:Root).Status | Should -Be 'Missing'
    }
    It 'flags a stale plan and keeps it usable' {
        New-HandoffFile -Path (Get-SuiteHandoffPath -SiteCode 'AB1' -Kind 'ring-plan' -DataRoot $script:Root) -Created ([datetime]::UtcNow.AddDays(-5))
        $r = Get-DhHandoffPlan -SiteFolder 'AB1' -DataRoot $script:Root -StaleAfterHours 72
        $r.Status | Should -Be 'Ok'
        $r.IsStale | Should -BeTrue
        $r.AgeText | Should -Be '5 days old'
        (Get-DhHandoffPlan -SiteFolder 'AB1' -DataRoot $script:Root -StaleAfterHours 240).IsStale | Should -BeFalse
    }
    It 'flags a plan whose site code is another site' {
        New-HandoffFile -Path (Get-SuiteHandoffPath -SiteCode 'AB1' -Kind 'ring-plan' -DataRoot $script:Root) -Site 'CD2'
        (Get-DhHandoffPlan -SiteFolder 'AB1' -DataRoot $script:Root).SiteMismatch | Should -BeTrue
    }
    It 'refuses a plan with no header, so its age is not trusted' {
        New-HandoffFile -Path (Get-SuiteHandoffPath -SiteCode 'AB1' -Kind 'ring-plan' -DataRoot $script:Root) -NoHeader
        $r = Get-DhHandoffPlan -SiteFolder 'AB1' -DataRoot $script:Root
        $r.Status | Should -Be 'NoDate'
        $r.Message | Should -Match 'age is not known'
    }
    It 'refuses a plan with a newer schema version and names the file' {
        New-HandoffFile -Path (Get-SuiteHandoffPath -SiteCode 'AB1' -Kind 'ring-plan' -DataRoot $script:Root) -Version 2
        $r = Get-DhHandoffPlan -SiteFolder 'AB1' -DataRoot $script:Root
        $r.Status | Should -Be 'TooNew'
        $r.Message | Should -Match 'ring-plan.json'
        $r.Message | Should -Match 'version 2'
    }
    It 'refuses a malformed file' {
        $p = Get-SuiteHandoffPath -SiteCode 'AB1' -Kind 'ring-plan' -DataRoot $script:Root
        [void](New-Item -ItemType Directory -Path (Split-Path $p -Parent) -Force)
        [IO.File]::WriteAllText($p, '{ "Format": "SuiteRingPlan", "Rings": [')
        (Get-DhHandoffPlan -SiteFolder 'AB1' -DataRoot $script:Root).Status | Should -Be 'Malformed'
    }
    It 'refuses a plan that the ring plan loader refuses (a built-in collection)' {
        $p = Get-SuiteHandoffPath -SiteCode 'AB1' -Kind 'ring-plan' -DataRoot $script:Root
        New-HandoffFile -Path $p
        $doc = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json
        $doc.Rings[0].CollectionID = 'SMS00001'
        [IO.File]::WriteAllText($p, ($doc | ConvertTo-Json -Depth 8))
        $r = Get-DhHandoffPlan -SiteFolder 'AB1' -DataRoot $script:Root
        $r.Status | Should -Be 'Invalid'
        $r.Message | Should -Not -BeNullOrEmpty
    }
}

Describe 'Get-RingPlanUpdate' {
    It 'lists the updates of every ring as read-only rows with the contract fields' {
        $rows = @(Get-RingPlanUpdate -Path $script:Fixture)
        $rows.Count | Should -BeGreaterThan 0
        foreach ($n in 'Wave', 'Sequence', 'ArticleId', 'Title', 'SeverityName', 'NeededDevices', 'CIId') { $rows[0].PSObject.Properties.Name | Should -Contain $n }
        $rows[0].Wave | Should -Be 'Wave 1 - Zone 01'
        $rows[0].Sequence | Should -Be 1
        $rows[0].ArticleId | Should -Be '5012170'
        $rows[0].CIId | Should -Be 16777772
    }
    It 'gives no rows for a plan with no Updates, a missing file or bad text' {
        $p = Join-Path $TestDrive 'plain.json'
        [IO.File]::WriteAllText($p, '{ "Name": "x", "Rings": [ { "Name": "r" } ] }')
        @(Get-RingPlanUpdate -Path $p).Count | Should -Be 0
        @(Get-RingPlanUpdate -Path (Join-Path $TestDrive 'none.json')).Count | Should -Be 0
        [IO.File]::WriteAllText($p, 'nonsense')
        @(Get-RingPlanUpdate -Path $p).Count | Should -Be 0
    }
}

Describe 'Copy-DhHandoffPlan' {
    BeforeEach {
        $script:Root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $script:Dir = Join-Path $script:Root 'rings'
        New-HandoffFile -Path (Get-SuiteHandoffPath -SiteCode 'AB1' -Kind 'ring-plan' -DataRoot $script:Root)
        $script:H = Get-DhHandoffPlan -SiteFolder 'AB1' -DataRoot $script:Root
    }
    It 'copies the plan into the plan folder and leaves no temp file' {
        $r = Copy-DhHandoffPlan -Handoff $script:H -RingPlanDir $script:Dir
        $r.Status | Should -Be 'Imported'
        $r.Path | Should -Be (Join-Path $script:Dir 'UpdatePlanner-a1b2c3d4.json')
        (Get-FileHash $r.Path).Hash | Should -Be (Get-FileHash $script:H.Path).Hash
        @(Get-ChildItem $script:Dir -Force).Count | Should -Be 1
        (Import-RingPlan -Path $r.Path).Rings.Count | Should -Be 3
    }
    It 'says AlreadyImported for the same content and writes nothing' {
        [void](Copy-DhHandoffPlan -Handoff $script:H -RingPlanDir $script:Dir)
        (Copy-DhHandoffPlan -Handoff $script:H -RingPlanDir $script:Dir).Status | Should -Be 'AlreadyImported'
    }
    It 'does not replace a file with other content unless asked' {
        $target = Join-Path $script:Dir 'UpdatePlanner-a1b2c3d4.json'
        [void](New-Item -ItemType Directory -Path $script:Dir -Force)
        [IO.File]::WriteAllText($target, 'edited by hand')
        (Copy-DhHandoffPlan -Handoff $script:H -RingPlanDir $script:Dir).Status | Should -Be 'Exists'
        Get-Content $target -Raw | Should -Be 'edited by hand'
        (Copy-DhHandoffPlan -Handoff $script:H -RingPlanDir $script:Dir -Overwrite).Status | Should -Be 'Imported'
        (Get-FileHash $target).Hash | Should -Be (Get-FileHash $script:H.Path).Hash
    }
    It 'CheckOnly writes nothing' {
        (Copy-DhHandoffPlan -Handoff $script:H -RingPlanDir $script:Dir -CheckOnly).Status | Should -Be 'Ready'
        @(Get-ChildItem $script:Dir -Force -ErrorAction SilentlyContinue).Count | Should -Be 0
    }
    It 'copies nothing from a plan that was refused' {
        $bad = [pscustomobject]@{ Status = 'TooNew'; Message = 'newer'; Path = ''; PlanName = 'x' }
        $r = Copy-DhHandoffPlan -Handoff $bad -RingPlanDir $script:Dir
        $r.Status | Should -Be 'Failed'
        $r.Message | Should -Be 'newer'
    }
    It 'turns a plan name into a safe file name' {
        $h = $script:H.PSObject.Copy()
        $h.PlanName = '..\evil:name'
        $r = Copy-DhHandoffPlan -Handoff $h -RingPlanDir $script:Dir
        (Split-Path $r.Path -Parent) | Should -Be $script:Dir
        (Split-Path $r.Path -Leaf) | Should -Be '.._evil_name.json'
    }
}

Describe 'Source text' {
    It 'shows the tool, the site, the age and the stale mark of a plan' {
        $h = [pscustomobject]@{ Status = 'Ok'; ProducerName = 'UpdatePlanner'; ProducerVersion = '1.2'; SiteCode = 'AB1'; AgeText = '5 days old'; IsStale = $true; SiteMismatch = $false; Message = 'Ok' }
        $t = Format-DhHandoffText -Handoff $h -StaleLimitText '3 days'
        $t | Should -Be 'Source: UpdatePlanner 1.2, site AB1, 5 days old. STALE: older than 3 days.'
        $h.IsStale = $false; $h.SiteMismatch = $true
        Format-DhHandoffText -Handoff $h | Should -Be 'Source: UpdatePlanner 1.2, site AB1, 5 days old. The plan is for another site than this window.'
    }
    It 'returns the message of a refused file' {
        Format-DhHandoffText -Handoff ([pscustomobject]@{ Status = 'Malformed'; Message = 'bad file' }) | Should -Be 'bad file'
    }
}

Describe 'Import-RolloutCheckResult: source and age' {
    BeforeAll {
        function Write-CheckCsv {
            param([string]$Path, [string[]]$Lines)
            $header = '"Application","Vendor","CMName","InstalledSoftwareName","Architecture","ConfigMgrVersion","LatestVersion","Status","Source","CheckedAtUtc"'
            $dir = Split-Path $Path -Parent
            if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
            Set-Content -LiteralPath $Path -Value (@($header) + $Lines) -Encoding UTF8
        }
        function Get-Stamp { param([double]$HoursAgo) return [datetime]::UtcNow.AddHours(-$HoursAgo).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'") }
    }

    It 'lists the file, the tool, the age of the newest row and the rows taken' {
        $f = Join-Path $TestDrive 'old-place\CheckConfigMgrResult.csv'
        Write-CheckCsv -Path $f -Lines @(('"X","","X","X*","x64","1.0","2.0","UpdateAvailable","APD","{0}"' -f (Get-Stamp 5)), ('"Y","","Y","Y*","x64","1.0","1.0","UpToDate","APD","{0}"' -f (Get-Stamp 1)))
        $r = Import-RolloutCheckResult -Path $f -StaleAfterHours 72
        @($r.Sources).Count | Should -Be 1
        $r.Sources[0].Path | Should -Be $f
        $r.Sources[0].Source | Should -Be 'APD'
        $r.Sources[0].SiteCode | Should -Be ''
        $r.Sources[0].AgeText | Should -Be '1 hour old'
        $r.Sources[0].IsStale | Should -BeFalse
        $r.Sources[0].Rows | Should -Be 1
    }
    It 'reads the site from the site folder of the path' {
        $f = Join-Path $TestDrive 'data\sites\AB1\results\CheckConfigMgrResult.csv'
        Write-CheckCsv -Path $f -Lines @(('"X","","X","X*","x64","1.0","2.0","UpdateAvailable","APD","{0}"' -f (Get-Stamp 2)))
        $r = Import-RolloutCheckResult -Path $f -ActiveSite 'AB1'
        $r.Sources[0].SiteCode | Should -Be 'AB1'
        $r.Sources[0].SiteMismatch | Should -BeFalse
        (Import-RolloutCheckResult -Path $f -ActiveSite 'CD2').Sources[0].SiteMismatch | Should -BeTrue
    }
    It 'flags a stale file, and still returns its rows' {
        $f = Join-Path $TestDrive 'stale\CheckConfigMgrResult.csv'
        Write-CheckCsv -Path $f -Lines @(('"X","","X","X*","x64","1.0","2.0","UpdateAvailable","AppPackager","{0}"' -f (Get-Stamp 120)))
        $r = Import-RolloutCheckResult -Path $f -StaleAfterHours 72
        @($r.Rows).Count | Should -Be 1
        $r.Sources[0].IsStale | Should -BeTrue
        $r.Sources[0].AgeText | Should -Be '5 days old'
        (Import-RolloutCheckResult -Path $f -StaleAfterHours 200).Sources[0].IsStale | Should -BeFalse
    }
    It 'treats a file with no valid time as stale with an unknown age' {
        $f = Join-Path $TestDrive 'notime\CheckConfigMgrResult.csv'
        Write-CheckCsv -Path $f -Lines @('"X","","X","X*","x64","1.0","2.0","UpdateAvailable","APD","not a time"')
        $r = Import-RolloutCheckResult -Path $f
        $r.Sources[0].IsStale | Should -BeTrue
        $r.Sources[0].AgeText | Should -Be ''
        (Format-RolloutSourceText -Source $r.Sources[0]) | Should -Match 'age not known'
        (Format-RolloutSourceText -Source $r.Sources[0]) | Should -Match 'STALE'
    }
    It 'lists each file that gave rows when the apps disagree, and drops a file that gave none' {
        $a = Join-Path $TestDrive 'm1\CheckConfigMgrResult.csv'; $b = Join-Path $TestDrive 'm2\CheckConfigMgrResult.csv'; $c = Join-Path $TestDrive 'm3\CheckConfigMgrResult.csv'
        Write-CheckCsv -Path $a -Lines @(('"X","","X","X*","x64","1.0","2.0","UpdateAvailable","APD","{0}"' -f (Get-Stamp 30)))
        Write-CheckCsv -Path $b -Lines @(('"X","","X","X*","x64","1.0","2.5","UpdateAvailable","AppPackager","{0}"' -f (Get-Stamp 3)))
        Write-CheckCsv -Path $c -Lines @(('"Z","","Z","Z*","x64","1.0","1.0","UpToDate","APD","{0}"' -f (Get-Stamp 3)))
        $r = Import-RolloutCheckResult -Path $a, $b, $c
        $r.Rows[0].NewVersion | Should -Be '2.5'
        @($r.Sources).Count | Should -Be 1
        $r.Sources[0].Source | Should -Be 'AppPackager'
    }
    It 'formats the source text with the tool, the site, the file and the age' {
        $s = [pscustomobject]@{ Source = 'APD'; SiteCode = 'AB1'; Path = 'C:\x\CheckConfigMgrResult.csv'; AgeText = '2 days old'; IsStale = $true; SiteMismatch = $false }
        Format-RolloutSourceText -Source $s | Should -Be 'APD, site AB1, C:\x\CheckConfigMgrResult.csv, 2 days old - STALE'
    }
    It 'returns no source for a missing file' {
        @((Import-RolloutCheckResult -Path (Join-Path $TestDrive 'nofile.csv')).Sources).Count | Should -Be 0
    }
}

Describe 'Get-RolloutCheckResultPath' {
    It 'puts the results folder of the site first' {
        $p = @(Get-RolloutCheckResultPath -SiteFolder 'AB1' -DataRoot 'C:\data')
        $p[0] | Should -Be 'C:\data\sites\AB1\results\CheckConfigMgrResult.csv'
        $p.Count | Should -BeGreaterThan 1
    }
    It 'keeps the old places when no site is given' {
        @(Get-RolloutCheckResultPath) | ForEach-Object { $_ | Should -Not -Match 'sites\\' }
    }
}
