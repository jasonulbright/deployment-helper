#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }

<#
.SYNOPSIS
    Pester 5 tests for the update rollout functions of the DeploymentHelperCommon module.

.DESCRIPTION
    Unit tests with stubbed ConfigurationManager cmdlets. The version tests check the LIKE patterns of the query
    against a numeric version compare for every combination of a set of components.
    Run: Invoke-Pester -Path .\Tests\UpdateRollout.Tests.ps1 -Output Detailed
#>

BeforeAll {
    . (Join-Path $PSScriptRoot 'CMStubs.ps1')
    Set-CMTestStub -Name 'Get-CMApplication', 'Get-CMCollection', 'Get-CMApplicationDeployment', 'Invoke-CMWmiQuery', 'New-CMSchedule',
        'New-CMDeviceCollection', 'Add-CMDeviceCollectionQueryMembershipRule', 'Move-CMObject', 'Invoke-CMCollectionUpdate',
        'Remove-CMCollection', 'Get-CMDeploymentType', 'New-CMApplicationDeployment'

    Import-Module (Join-Path $PSScriptRoot '..\Module\DeploymentHelperCommon.psd1') -Force
    Initialize-Logging -LogPath (Join-Path $TestDrive 'rollout.log')

    # The LIKE operator of WQL as a regular expression: % is any text, a bracket range is a character class.
    function ConvertTo-TestLikeRegex {
        param([string[]]$Pattern)
        $alts = foreach ($p in $Pattern) { '(?:' + ([regex]::Escape($p) -replace '\\\[', '[' -replace '\\\]', ']' -replace '%', '.*') + ')' }
        return [regex]::new('^(?:' + ($alts -join '|') + ')$', 'Compiled')
    }

    function New-TestItem {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Test helper; builds test data.')]
        param([string]$Stage = 'Prod', [string]$Role = 'Server', [string]$Name = 'App - Server - Installed - Version LT 2.0', [bool]$UseExisting = $false, [string]$Limit = 'All Systems')
        [PSCustomObject]@{
            Index = 1; Key = "App|2.0|$Role"; Stage = $Stage; Application = 'App'; Version = '2.0'; SoftwareName = 'App'
            Role = $Role; CollectionName = $Name; CollectionId = ''; UseExisting = $UseExisting; LimitingName = $Limit
            Query = 'select 1'; Exists = $false; MemberCount = $null; Status = 'Planned'; Message = ''; DeploymentId = ''
        }
    }
}

Describe 'ConvertTo-RolloutVersionPart' {
    It 'splits <Text>' -ForEach @(
        @{ Text = '24.8'; Expected = '24.8' }
        @{ Text = 'v1.2.3.4'; Expected = '1.2.3.4' }
        @{ Text = '007.010'; Expected = '7.10' }
        @{ Text = ' 5 '; Expected = '5' }
    ) {
        (ConvertTo-RolloutVersionPart -Version $Text) -join '.' | Should -Be $Expected
    }

    It 'returns $null for <Text>' -ForEach @(
        @{ Text = '' }, @{ Text = '1.2.3-beta' }, @{ Text = '1..2' }, @{ Text = 'abc' }, @{ Text = '1.2.3.4.5.6.7' }, @{ Text = '1.2a' }
    ) {
        ConvertTo-RolloutVersionPart -Version $Text | Should -BeNullOrEmpty
    }
}

Describe 'Compare-RolloutVersion' {
    It 'compares <Left> with <Right> as <Expected>' -ForEach @(
        @{ Left = '9.0'; Right = '10.0'; Expected = -1 }
        @{ Left = '10.0'; Right = '9.0'; Expected = 1 }
        @{ Left = '24.8'; Right = '24.08'; Expected = 0 }
        @{ Left = '1.2'; Right = '1.2.0.0'; Expected = 0 }
        @{ Left = '1.2'; Right = '1.2.0.1'; Expected = -1 }
        @{ Left = '2.0.0.1'; Right = '2'; Expected = 1 }
        @{ Left = '26.03.00.0'; Right = '26.06'; Expected = -1 }
    ) {
        Compare-RolloutVersion -Left $Left -Right $Right | Should -Be $Expected
    }

    It 'throws for text that is not a version' {
        { Compare-RolloutVersion -Left '1.2-beta' -Right '1.2' } | Should -Throw
    }
}

Describe 'Get-RolloutVersionLikeClause' {
    It 'matches exactly the versions that the numeric compare puts below <Threshold>' -ForEach @(
        @{ Threshold = '24.8' }, @{ Threshold = '1' }, @{ Threshold = '10' }, @{ Threshold = '2.0.1' }, @{ Threshold = '1.2.3.4' },
        @{ Threshold = '100.0' }, @{ Threshold = '9.99' }, @{ Threshold = '0.5' }, @{ Threshold = '120.0.6099.109' }, @{ Threshold = '5.0.0.0' }
    ) {
        $rx = ConvertTo-TestLikeRegex -Pattern (Get-RolloutVersionLikeClause -Version $Threshold)
        $values = '0', '1', '2', '5', '9', '10', '11', '19', '20', '99', '100', '101', '109', '120'
        $bad = New-Object System.Collections.Generic.List[string]
        $texts = New-Object System.Collections.Generic.List[string]
        foreach ($a in $values) {
            $texts.Add($a)
            foreach ($b in $values) {
                $texts.Add("$a.$b")
                if ($values.IndexOf($b) % 3 -eq 0) { foreach ($c in $values) { $texts.Add("$a.$b.$c"); if ($values.IndexOf($c) % 4 -eq 0) { foreach ($d in $values) { $texts.Add("$a.$b.$c.$d") } } } }
            }
        }
        # The same versions written with one leading zero in a component (7-Zip writes 26.03.00.0 for 26.3).
        foreach ($t in @($texts | Where-Object { $_ -match '^\d+(\.\d+){1,3}$' } | Select-Object -First 1500)) {
            $c = $t.Split('.')
            $texts.Add((($c | ForEach-Object { '0' + $_ }) -join '.'))
            $texts.Add((@($c[0]) + @($c | Select-Object -Skip 1 | ForEach-Object { '0' + $_ }) -join '.'))
        }
        foreach ($t in $texts) {
            $below = ((Compare-RolloutVersion -Left $t -Right $Threshold) -lt 0)
            if ($rx.IsMatch($t) -ne $below) { $bad.Add($t) }
        }
        $texts.Count | Should -BeGreaterThan 1000
        ($bad | Select-Object -First 5) -join ', ' | Should -Be ''
    }

    It 'never matches a text whose first components are not digits' {
        $rx = ConvertTo-TestLikeRegex -Pattern (Get-RolloutVersionLikeClause -Version '30.0')
        foreach ($t in '2024a', 'v1.2', '1a.2', 'a.1', '.5', '', '1 2') { $rx.IsMatch($t) | Should -BeFalse -Because $t }
        $rx.IsMatch('1.2.3-beta') | Should -BeTrue -Because 'the first component 1 is below 30, whatever follows'
        $rx.IsMatch('31.0-beta') | Should -BeFalse
    }

    It 'reads a component with a leading zero as its number (7-Zip 26.03.00.0 is below 26.06)' {
        $rx = ConvertTo-TestLikeRegex -Pattern (Get-RolloutVersionLikeClause -Version '26.06')
        $rx.IsMatch('26.03.00.0') | Should -BeTrue
        $rx.IsMatch('26.3.0.0') | Should -BeTrue
        $rx.IsMatch('26.06.00.0') | Should -BeFalse
        $rx.IsMatch('26.07.00.0') | Should -BeFalse
        $rx.IsMatch('25.12.00.0') | Should -BeTrue
    }

    It 'stays a usable query size for a four-part version' {
        @(Get-RolloutVersionLikeClause -Version '126.0.6478.127').Count | Should -BeLessThan 600
        (New-RolloutQuery -SoftwareName 'Google Chrome' -Version '126.0.6478.127').Length | Should -BeLessThan 60000
    }

    It 'compares the numbers of a component, not its text (9 is below 10)' {
        $rx = ConvertTo-TestLikeRegex -Pattern (Get-RolloutVersionLikeClause -Version '10.0')
        $rx.IsMatch('9.5') | Should -BeTrue
        $rx.IsMatch('10.0') | Should -BeFalse
        $rx.IsMatch('10.0.0.0') | Should -BeFalse
        $rx.IsMatch('100.0') | Should -BeFalse
    }

    It 'uses only LIKE syntax that WQL supports' {
        foreach ($p in (Get-RolloutVersionLikeClause -Version '126.0.6478.127')) { $p | Should -Match '^[0-9.\[\]%-]+$' }
    }

    It 'throws for a version that is not a dotted number' {
        { Get-RolloutVersionLikeClause -Version '1.2-beta' } | Should -Throw
    }
}

Describe 'New-RolloutQuery' {
    It 'searches both installed-software classes for the name and the version' {
        $q = New-RolloutQuery -SoftwareName '7-Zip' -Version '26.06'
        $q | Should -Match 'from SMS_R_System where'
        $q | Should -Match 'SMS_G_System_ADD_REMOVE_PROGRAMS\.DisplayName like ''%7-Zip%'''
        $q | Should -Match 'SMS_G_System_ADD_REMOVE_PROGRAMS_64\.DisplayName like ''%7-Zip%'''
        $q | Should -Match 'SMS_G_System_ADD_REMOVE_PROGRAMS_64\.Version like ''\[0-9\]'
    }

    It 'doubles a quote and escapes the LIKE wildcards of the name' {
        $q = New-RolloutQuery -SoftwareName "O'Neil_100%[x]" -Version '2'
        $q | Should -Match ([regex]::Escape("like '%O''Neil[_]100[%][[]x]%'"))
    }

    It 'adds the limiting collection only when it is given' {
        (New-RolloutQuery -SoftwareName 'A' -Version '2') | Should -Not -Match 'SMS_CM_RES_COLL'
        (New-RolloutQuery -SoftwareName 'A' -Version '2' -LimitingCollectionId 'sms00001') | Should -Match 'SMS_CM_RES_COLL_SMS00001\)$'
    }

    It 'refuses an empty name, a bad version and a bad limiting collection' {
        { New-RolloutQuery -SoftwareName ' ' -Version '2' } | Should -Throw
        { New-RolloutQuery -SoftwareName 'A' -Version 'x' } | Should -Throw
        { New-RolloutQuery -SoftwareName 'A' -Version '2' -LimitingCollectionId "x'; drop" } | Should -Throw
    }
}

Describe 'Expand-RolloutName and Test-RolloutName' {
    It 'fills the default template' {
        Expand-RolloutName -Template '{App} - {Role} - Installed - Version LT {Version}' -Application 'Notepad++' -Role 'Server' -Version '8.6.2' |
            Should -Be 'Notepad++ - Server - Installed - Version LT 8.6.2'
    }

    It 'fills a placeholder in any case and keeps a dollar sign of the application name' {
        Expand-RolloutName -Template '{app}/{ROLE}/{Version}' -Application 'A$1' -Role 'Test' -Version '1' | Should -Be 'A$1/Test/1'
    }

    It 'accepts a normal name' { Test-RolloutName -Name 'App - Server - Installed - Version LT 2.0' | Should -Be '' }
    It 'refuses <Why>' -ForEach @(
        @{ Why = 'an empty name'; Name = ' ' }
        @{ Why = 'a placeholder that is left'; Name = 'App {Foo}' }
        @{ Why = 'a name that is too long'; Name = ('x' * 256) }
    ) {
        Test-RolloutName -Name $Name | Should -Not -Be ''
    }
}

Describe 'ConvertTo-RolloutSettings' {
    It 'returns the defaults for empty preferences' {
        $s = ConvertTo-RolloutSettings -Prefs @{}
        $s.RolloutNameTemplate | Should -Be '{App} - {Role} - Installed - Version LT {Version}'
        $s.RolloutRefreshDays | Should -Be 1
        $s.RolloutDefaultRole | Should -Be 'Both'
        $s.RolloutNotification | Should -Be 'DisplayAll'
    }

    It 'takes saved values and drops invalid ones' {
        $s = ConvertTo-RolloutSettings -Prefs @{ RolloutLimitingServer = ' All Servers '; RolloutRefreshDays = '7'; RolloutProdDeadlineHours = 'abc'; RolloutNotification = 'Loud'; RolloutDefaultRole = 'Workstation' }
        $s.RolloutLimitingServer | Should -Be 'All Servers'
        $s.RolloutRefreshDays | Should -Be 7
        $s.RolloutProdDeadlineHours | Should -Be 72
        $s.RolloutNotification | Should -Be 'DisplayAll'
        $s.RolloutDefaultRole | Should -Be 'Workstation'
    }
}

Describe 'ConvertFrom-RolloutList' {
    It 'reads a CSV with a header' {
        $r = ConvertFrom-RolloutList -Text "Application,NewVersion,InstalledSoftwareName,Role,TestCollection`n7-Zip,26.06,7-Zip,Server,Test PCs`nNotepad++,v8.6.2,,Workstation,"
        $r.Errors | Should -BeNullOrEmpty
        @($r.Rows).Count | Should -Be 2
        $r.Rows[0].Application | Should -Be '7-Zip'
        $r.Rows[0].SoftwareName | Should -Be '7-Zip'
        $r.Rows[0].Role | Should -Be 'Server'
        $r.Rows[0].TestCollection | Should -Be 'Test PCs'
        $r.Rows[1].NewVersion | Should -Be 'v8.6.2'
        $r.Rows[1].Role | Should -Be 'Workstation'
    }

    It 'reads a semicolon CSV and a tab CSV' {
        (ConvertFrom-RolloutList -Text "Application;NewVersion`nA;1.0").Rows[0].NewVersion | Should -Be '1.0'
        (ConvertFrom-RolloutList -Text "Application`tNewVersion`nA`t2.5").Rows[0].NewVersion | Should -Be '2.5'
    }

    It 'reads a list without a header: separators and a trailing version' {
        $r = ConvertFrom-RolloutList -Text "7-Zip, 26.06`nAudacity;3.7.9`nGoogle Chrome 126.0.6478.127`nPuTTY`t0.81" -DefaultRole 'Server'
        $r.Errors | Should -BeNullOrEmpty
        @($r.Rows).Count | Should -Be 4
        $r.Rows[2].Application | Should -Be 'Google Chrome'
        $r.Rows[2].NewVersion | Should -Be '126.0.6478.127'
        $r.Rows[3].Role | Should -Be 'Server'
    }

    It 'names the line of every refused row and keeps the good rows' {
        $r = ConvertFrom-RolloutList -Text "Application,NewVersion,Role`nGood,1.0,Server`nBad,1.0-beta,Server`n,2.0,Server`nWorse,1.0,Cloud"
        @($r.Rows).Count | Should -Be 1
        @($r.Errors).Count | Should -Be 3
        $r.Errors[0] | Should -Match 'Line 3'
    }

    It 'refuses an empty text and a header without a version column' {
        (ConvertFrom-RolloutList -Text '').Errors | Should -Not -BeNullOrEmpty
        (ConvertFrom-RolloutList -Text "Application,Role`nA,Server").Errors[0] | Should -Match 'NewVersion'
    }
}

Describe 'New-RolloutPlan' {
    BeforeAll { $script:Settings = ConvertTo-RolloutSettings -Prefs @{ RolloutLimitingServer = 'All Servers'; RolloutLimitingWorkstation = 'All Workstations'; RolloutLimitingTest = 'Test Devices' } }

    It 'makes a test item and a production item per role for the role Both' {
        $plan = New-RolloutPlan -Settings $script:Settings -Rows @((New-RolloutRow -Application '7-Zip' -NewVersion '26.06' -SoftwareName '7-Zip' -Role 'Both'))
        $plan.Errors | Should -BeNullOrEmpty
        @($plan.Items).Count | Should -Be 3
        ($plan.Items | ForEach-Object { '{0}:{1}' -f $_.Stage, $_.Role }) -join ',' | Should -Be 'Test:Test,Prod:Server,Prod:Workstation'
        $plan.Items[1].CollectionName | Should -Be '7-Zip - Server - Installed - Version LT 26.06'
        $plan.Items[1].LimitingName | Should -Be 'All Servers'
        $plan.Items[2].LimitingName | Should -Be 'All Workstations'
        $plan.Items[0].CollectionName | Should -Be '7-Zip - Test - Installed - Version LT 26.06'
        $plan.Items[0].LimitingName | Should -Be 'Test Devices'
        $plan.Items[0].Query | Should -Match 'ADD_REMOVE_PROGRAMS'
        ($plan.Items | ForEach-Object { $_.Index }) -join ',' | Should -Be '1,2,3'
    }

    It 'makes one production item for Server or Workstation' {
        $p = New-RolloutPlan -Settings $script:Settings -Rows @((New-RolloutRow -Application 'A' -NewVersion '2' -SoftwareName 'A' -Role 'Workstation'))
        ($p.Items | Where-Object Stage -eq 'Prod' | ForEach-Object Role) | Should -Be 'Workstation'
    }

    It 'uses an existing test collection of the row, and of the setting, and creates no query for it' {
        $row = New-RolloutRow -Application 'A' -NewVersion '2' -SoftwareName 'A' -Role 'Server' -TestCollection 'My Test PCs'
        $p = New-RolloutPlan -Settings $script:Settings -Rows @($row)
        $p.Items[0].UseExisting | Should -BeTrue
        $p.Items[0].CollectionName | Should -Be 'My Test PCs'
        $p.Items[0].Query | Should -Be ''
        $s2 = ConvertTo-RolloutSettings -Prefs @{ RolloutTestCollection = 'Default Test' }
        (New-RolloutPlan -Settings $s2 -Rows @((New-RolloutRow -Application 'A' -NewVersion '2' -SoftwareName 'A' -Role 'Server'))).Items[0].CollectionName | Should -Be 'Default Test'
    }

    It 'reports a bad version, a missing software name, a bad role and an empty limiting collection' {
        $rows = @(
            [PSCustomObject]@{ Application = 'A'; NewVersion = '1.x'; SoftwareName = 'A'; Role = 'Server'; TestCollection = '' }
            [PSCustomObject]@{ Application = 'B'; NewVersion = '1.0'; SoftwareName = ''; Role = 'Server'; TestCollection = '' }
            [PSCustomObject]@{ Application = 'C'; NewVersion = '1.0'; SoftwareName = 'C'; Role = 'Cloud'; TestCollection = '' }
        )
        $p = New-RolloutPlan -Settings $script:Settings -Rows $rows
        @($p.Errors).Count | Should -Be 3
        $s3 = ConvertTo-RolloutSettings -Prefs @{ RolloutLimitingServer = ''; RolloutLimitingTest = 'x' }
        $s3.RolloutLimitingServer = ''
        (New-RolloutPlan -Settings $s3 -Rows @((New-RolloutRow -Application 'D' -NewVersion '1' -SoftwareName 'D' -Role 'Server'))).Errors | Should -Match 'limiting collection is empty'
    }

    It 'reports two items with the same collection name' {
        $r = New-RolloutRow -Application 'A' -NewVersion '2' -SoftwareName 'A' -Role 'Server'
        (New-RolloutPlan -Settings $script:Settings -Rows @($r, $r)).Errors | Should -Match 'Two items'
    }

    It 'lets two applications share one existing test collection' {
        $s = ConvertTo-RolloutSettings -Prefs @{ RolloutTestCollection = 'Shared Test' }
        $rows = @((New-RolloutRow -Application 'A' -NewVersion '2' -SoftwareName 'A' -Role 'Server'), (New-RolloutRow -Application 'B' -NewVersion '3' -SoftwareName 'B' -Role 'Server'))
        $p = New-RolloutPlan -Settings $s -Rows $rows
        $p.Errors | Should -BeNullOrEmpty
        @($p.Items | Where-Object { $_.Stage -eq 'Test' }).Count | Should -Be 2
    }

    It 'reports a name template that leaves a placeholder' {
        $s = ConvertTo-RolloutSettings -Prefs @{ RolloutNameTemplate = '{App} {Foo} {Version}' }
        (New-RolloutPlan -Settings $s -Rows @((New-RolloutRow -Application 'A' -NewVersion '2' -SoftwareName 'A' -Role 'Server'))).Errors | Should -Match 'placeholder'
    }

    It 'writes the version without a leading v and keeps the numbers as typed' {
        $p = New-RolloutPlan -Settings $script:Settings -Rows @((New-RolloutRow -Application 'A' -NewVersion 'v8.6.2' -SoftwareName 'A' -Role 'Server'))
        $p.Items[1].CollectionName | Should -Be 'A - Server - Installed - Version LT 8.6.2'
    }
}

Describe 'Get-RolloutSchedule and Test-RolloutSchedule' {
    It 'counts the deadline from the available time' {
        $s = ConvertTo-RolloutSettings -Prefs @{ RolloutAvailableDelayHours = 2; RolloutTestDeadlineHours = 24; RolloutProdDeadlineHours = 72 }
        $now = [datetime]'2026-11-21T08:00:00'
        $t = Get-RolloutSchedule -Stage Test -Settings $s -Now $now
        $t.Available | Should -Be ([datetime]'2026-11-21T10:00:00')
        $t.Deadline | Should -Be ([datetime]'2026-11-22T10:00:00')
        (Get-RolloutSchedule -Stage Prod -Settings $s -Now $now).Deadline | Should -Be ([datetime]'2026-11-24T10:00:00')
    }

    It 'blocks a deadline that is not after the available time' {
        $s = ConvertTo-RolloutSettings -Prefs @{ RolloutTestDeadlineHours = 0 }
        Test-RolloutSchedule -Schedule (Get-RolloutSchedule -Stage Test -Settings $s) | Should -Match 'at least 1 hour'
        Test-RolloutSchedule -Schedule (Get-RolloutSchedule -Stage Prod -Settings (ConvertTo-RolloutSettings -Prefs @{})) | Should -Be ''
    }
}

Describe 'Get-RolloutSoftwareNameSuggestion' {
    It 'cuts the version from a display name' {
        Get-RolloutBaseName -Name '7-Zip 26.03 (x64 edition)' | Should -Be '7-Zip'
        Get-RolloutBaseName -Name 'Google Chrome' | Should -Be 'Google Chrome'
        Get-RolloutBaseName -Name 'Python 3.12.4 (64-bit)' | Should -Be 'Python'
    }

    It 'suggests the most frequent inventory name, with the version cut' {
        Mock Invoke-CMWmiQuery -ModuleName DeploymentHelperCommon {
            @([PSCustomObject]@{ DisplayName = '7-Zip 26.03 (x64 edition)' }, [PSCustomObject]@{ DisplayName = '7-Zip 26.03 (x64 edition)' }, [PSCustomObject]@{ DisplayName = '7-Zip Helper' })
        }
        $r = Get-RolloutSoftwareNameSuggestion -ApplicationName '7-Zip 26.06 (x64 edition)'
        $r.Name | Should -Be '7-Zip'
        $r.Source | Should -Be 'Inventory'
    }

    It 'falls back to the application name when the inventory has no hit or the read fails' {
        Mock Invoke-CMWmiQuery -ModuleName DeploymentHelperCommon { @() }
        (Get-RolloutSoftwareNameSuggestion -ApplicationName 'Audacity 3.7.9').Name | Should -Be 'Audacity'
        Mock Invoke-CMWmiQuery -ModuleName DeploymentHelperCommon { throw 'boom' }
        (Get-RolloutSoftwareNameSuggestion -ApplicationName 'Audacity 3.7.9').Source | Should -Be 'Application name'
    }

    It 'looks the product code up first when it is given' {
        Mock Invoke-CMWmiQuery -ModuleName DeploymentHelperCommon { @([PSCustomObject]@{ DisplayName = 'Exact Name 1.0' }) }
        $r = Get-RolloutSoftwareNameSuggestion -ApplicationName 'Some App' -ProductCode '{12345678-1234-1234-1234-123456789ABC}'
        $r.Source | Should -Be 'Product code'
        $r.Name | Should -Be 'Exact Name'
    }
}

Describe 'Get-RolloutMemberPreview' {
    It 'counts the rows of the query inside the limiting collection' {
        Mock Invoke-CMWmiQuery -ModuleName DeploymentHelperCommon { @(1, 2, 3) }
        $r = Get-RolloutMemberPreview -Query 'select x from SMS_R_System where (a)' -LimitingCollectionId 'SMS00001'
        $r.Count | Should -Be 3
        Should -Invoke Invoke-CMWmiQuery -ModuleName DeploymentHelperCommon -ParameterFilter { $Query -match 'SMS_CM_RES_COLL_SMS00001\)$' } -Times 1
    }

    It 'returns a user message and no count when the query fails' {
        Mock Invoke-CMWmiQuery -ModuleName DeploymentHelperCommon { throw 'Invalid query' }
        $r = Get-RolloutMemberPreview -Query 'select x' -LimitingCollectionId 'SMS00001'
        $r.Count | Should -BeNullOrEmpty
        $r.Error | Should -Not -Match 'Invalid query'
    }
}

Describe 'Test-RolloutPlanOnSite' {
    BeforeEach {
        $script:App = [PSCustomObject]@{ LocalizedDisplayName = 'App'; SoftwareVersion = '2.0'; PackageID = 'MCM00001' }
        Mock Test-ApplicationExists -ModuleName DeploymentHelperCommon { $script:App }
        Mock Test-ContentDistributed -ModuleName DeploymentHelperCommon { @{ IsFullyDistributed = $true; NumberSuccess = 1; Targeted = 1 } }
        Mock Get-CMConnectionInfo -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ SiteCode = 'MCM'; SMSProvider = 'cm01' } }
        Mock Test-Path -ModuleName DeploymentHelperCommon { $true }
        Mock Test-CollectionValid -ModuleName DeploymentHelperCommon {
            if ($CollectionName -eq 'All Systems') { return [PSCustomObject]@{ Name = 'All Systems'; CollectionID = 'SMS00001'; CollectionType = 2; MemberCount = 9 } }
            return $null
        }
        Mock Test-DuplicateDeployment -ModuleName DeploymentHelperCommon { $null }
        Mock Invoke-CMWmiQuery -ModuleName DeploymentHelperCommon { @(1, 2) }
        $script:S = ConvertTo-RolloutSettings -Prefs @{}
    }

    It 'passes an item that creates a collection and previews its members' {
        $items = @((New-RolloutPlan -Settings $script:S -Rows @((New-RolloutRow -Application 'App' -NewVersion '2.0' -SoftwareName 'App' -Role 'Server'))).Items | Where-Object Stage -eq 'Prod')
        $r = Test-RolloutPlanOnSite -Items $items -Settings $script:S
        $r.Ok | Should -BeTrue
        $items[0].Status | Should -Be 'Ready'
        $items[0].MemberCount | Should -Be 2
        $items[0].Exists | Should -BeFalse
    }

    It 'reuses a collection that exists and creates nothing for it' {
        Mock Test-CollectionValid -ModuleName DeploymentHelperCommon {
            if ($CollectionName -eq 'All Systems') { return [PSCustomObject]@{ Name = 'All Systems'; CollectionID = 'SMS00001'; CollectionType = 2; MemberCount = 9 } }
            if ($CollectionName -eq 'App - Server - Installed - Version LT 2.0') { return [PSCustomObject]@{ Name = $CollectionName; CollectionID = 'MCM00099'; CollectionType = 2; MemberCount = 4 } }
            return $null
        }
        $items = @((New-RolloutPlan -Settings $script:S -Rows @((New-RolloutRow -Application 'App' -NewVersion '2.0' -SoftwareName 'App' -Role 'Server'))).Items | Where-Object Stage -eq 'Prod')
        $r = Test-RolloutPlanOnSite -Items $items -Settings $script:S
        $r.Ok | Should -BeTrue
        $items[0].Status | Should -Be 'Exists'
        $items[0].CollectionId | Should -Be 'MCM00099'
        $items[0].MemberCount | Should -Be 4
        $items[0].Message | Should -Match 'nothing is created'
        Should -Invoke Invoke-CMWmiQuery -ModuleName DeploymentHelperCommon -Times 0
    }

    It 'blocks an existing collection that has a deployment of the application already' {
        Mock Test-CollectionValid -ModuleName DeploymentHelperCommon {
            if ($CollectionName -eq 'All Systems') { return [PSCustomObject]@{ Name = 'All Systems'; CollectionID = 'SMS00001'; CollectionType = 2; MemberCount = 9 } }
            return [PSCustomObject]@{ Name = $CollectionName; CollectionID = 'MCM00099'; CollectionType = 2; MemberCount = 4 }
        }
        Mock Test-DuplicateDeployment -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ AssignmentID = 1 } }
        $items = @((New-RolloutPlan -Settings $script:S -Rows @((New-RolloutRow -Application 'App' -NewVersion '2.0' -SoftwareName 'App' -Role 'Server'))).Items | Where-Object Stage -eq 'Prod')
        $r = Test-RolloutPlanOnSite -Items $items -Settings $script:S
        $r.Ok | Should -BeFalse
        $items[0].Message | Should -Match 'already exists'
    }

    It 'blocks an existing collection with a built-in ID (SMS prefix)' {
        Mock Test-CollectionValid -ModuleName DeploymentHelperCommon {
            return [PSCustomObject]@{ Name = 'All Systems'; CollectionID = 'SMS00001'; CollectionType = 2; MemberCount = 9 }
        }
        $row = New-RolloutRow -Application 'App' -NewVersion '2.0' -SoftwareName 'App' -Role 'Server' -TestCollection 'All Systems'
        $items = @((New-RolloutPlan -Settings $script:S -Rows @($row)).Items | Where-Object Stage -eq 'Test')
        $r = Test-RolloutPlanOnSite -Items $items -Settings $script:S
        $r.Ok | Should -BeFalse
        $items[0].Message | Should -Match 'Built-in'
    }

    It 'blocks when the application is missing, its content is not distributed, the limiting collection or the folder is missing' {
        $items = @((New-RolloutPlan -Settings $script:S -Rows @((New-RolloutRow -Application 'App' -NewVersion '2.0' -SoftwareName 'App' -Role 'Server'))).Items | Where-Object Stage -eq 'Prod')
        Mock Test-ApplicationExists -ModuleName DeploymentHelperCommon { $null }
        (Test-RolloutPlanOnSite -Items $items -Settings $script:S).Ok | Should -BeFalse
        $items[0].Message | Should -Match 'was not found'

        Mock Test-ApplicationExists -ModuleName DeploymentHelperCommon { $script:App }
        Mock Test-ContentDistributed -ModuleName DeploymentHelperCommon { @{ IsFullyDistributed = $false; NumberSuccess = 0; Targeted = 2 } }
        (Test-RolloutPlanOnSite -Items $items -Settings $script:S).Ok | Should -BeFalse
        $items[0].Message | Should -Match 'not fully distributed'

        Mock Test-ContentDistributed -ModuleName DeploymentHelperCommon { @{ IsFullyDistributed = $true; NumberSuccess = 1; Targeted = 1 } }
        $s2 = ConvertTo-RolloutSettings -Prefs @{ RolloutLimitingServer = 'Nope' }
        $items2 = @((New-RolloutPlan -Settings $s2 -Rows @((New-RolloutRow -Application 'App' -NewVersion '2.0' -SoftwareName 'App' -Role 'Server'))).Items | Where-Object Stage -eq 'Prod')
        (Test-RolloutPlanOnSite -Items $items2 -Settings $s2).Ok | Should -BeFalse
        $items2[0].Message | Should -Match 'limiting collection'

        Mock Test-Path -ModuleName DeploymentHelperCommon { $false }
        $s3 = ConvertTo-RolloutSettings -Prefs @{ RolloutFolder = 'Updates\Missing' }
        $items3 = @((New-RolloutPlan -Settings $s3 -Rows @((New-RolloutRow -Application 'App' -NewVersion '2.0' -SoftwareName 'App' -Role 'Server'))).Items | Where-Object Stage -eq 'Prod')
        (Test-RolloutPlanOnSite -Items $items3 -Settings $s3).Ok | Should -BeFalse
        $items3[0].Message | Should -Match 'console folder'
    }

    It 'blocks when the member preview fails' {
        Mock Invoke-CMWmiQuery -ModuleName DeploymentHelperCommon { throw 'bad query' }
        $items = @((New-RolloutPlan -Settings $script:S -Rows @((New-RolloutRow -Application 'App' -NewVersion '2.0' -SoftwareName 'App' -Role 'Server'))).Items | Where-Object Stage -eq 'Prod')
        (Test-RolloutPlanOnSite -Items $items -Settings $script:S).Ok | Should -BeFalse
        $items[0].Status | Should -Be 'Blocked'
    }

    It 'warns, and does not block, when the application version differs from the row' {
        $script:App.SoftwareVersion = '1.9'
        $items = @((New-RolloutPlan -Settings $script:S -Rows @((New-RolloutRow -Application 'App' -NewVersion '2.0' -SoftwareName 'App' -Role 'Server'))).Items | Where-Object Stage -eq 'Prod')
        (Test-RolloutPlanOnSite -Items $items -Settings $script:S).Ok | Should -BeTrue
        $items[0].Message | Should -Match 'has version 1.9'
    }
}

Describe 'Invoke-RolloutItem' {
    BeforeEach {
        $script:Audit = Join-Path $TestDrive ('audit-{0}.jsonl' -f [guid]::NewGuid().ToString('N'))
        $script:S = ConvertTo-RolloutSettings -Prefs @{ RolloutFolder = 'Updates'; RolloutRefreshDays = 2 }
        $script:Sched = Get-RolloutSchedule -Stage Prod -Settings $script:S
        $script:App = [PSCustomObject]@{ LocalizedDisplayName = 'App'; SoftwareVersion = '2.0'; PackageID = 'MCM00001' }
        $script:NewCol = [PSCustomObject]@{ Name = 'App - Server - Installed - Version LT 2.0'; CollectionID = 'MCM00099'; CollectionType = 2; MemberCount = 0 }
        Mock Test-ApplicationExists -ModuleName DeploymentHelperCommon { $script:App }
        Mock Get-CMConnectionInfo -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ SiteCode = 'MCM'; SMSProvider = 'cm01' } }
        Mock Resolve-RolloutCollection -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ Name = 'All Systems'; CollectionID = 'SMS00001'; CollectionType = 2 } }
        Mock Test-CollectionValid -ModuleName DeploymentHelperCommon { if ($CollectionId) { return $script:NewCol } return $null }
        Mock Test-DuplicateDeployment -ModuleName DeploymentHelperCommon { $null }
        Mock New-CMSchedule -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ Sched = 1 } }
        Mock New-CMDeviceCollection -ModuleName DeploymentHelperCommon { $script:NewCol }
        Mock Add-CMDeviceCollectionQueryMembershipRule -ModuleName DeploymentHelperCommon { }
        Mock Move-CMObject -ModuleName DeploymentHelperCommon { }
        Mock Invoke-CMCollectionUpdate -ModuleName DeploymentHelperCommon { }
        Mock Remove-CMCollection -ModuleName DeploymentHelperCommon { }
        Mock New-CMApplicationDeployment -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ AssignmentID = 16777300; AssignmentUniqueID = '{G}' } }
    }

    It 'creates the collection with its schedule, rule and folder, deploys, and writes two audit records' {
        $item = New-TestItem
        $r = Invoke-RolloutItem -Item $item -Settings $script:S -Schedule $script:Sched -RunId 'RO-1' -AuditLogPath $script:Audit
        $r.Success | Should -BeTrue
        $r.Created | Should -BeTrue
        $r.CollectionId | Should -Be 'MCM00099'
        $r.DeploymentId | Should -Be '16777300'
        Should -Invoke New-CMDeviceCollection -ModuleName DeploymentHelperCommon -Times 1 -ParameterFilter { $LimitingCollectionId -eq 'SMS00001' -and $RefreshType -eq 'Periodic' -and $Name -eq 'App - Server - Installed - Version LT 2.0' }
        Should -Invoke New-CMSchedule -ModuleName DeploymentHelperCommon -Times 1 -ParameterFilter { $RecurInterval -eq 'Days' -and $RecurCount -eq 2 }
        Should -Invoke Add-CMDeviceCollectionQueryMembershipRule -ModuleName DeploymentHelperCommon -Times 1 -ParameterFilter { $CollectionId -eq 'MCM00099' -and $RuleName -eq 'Installed version below' -and $QueryExpression -eq 'select 1' }
        Should -Invoke Move-CMObject -ModuleName DeploymentHelperCommon -Times 1 -ParameterFilter { $FolderPath -eq 'MCM:\DeviceCollection\Updates' }
        Should -Invoke Invoke-CMCollectionUpdate -ModuleName DeploymentHelperCommon -Times 1
        Should -Invoke New-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 1 -ParameterFilter { $DeployPurpose -eq 'Required' -and $CollectionName -eq 'App - Server - Installed - Version LT 2.0' -and $DeadlineDateTime -gt $AvailableDateTime }
        $records = @(Get-Content -LiteralPath $script:Audit | ForEach-Object { $_ | ConvertFrom-Json })
        $records.Count | Should -Be 2
        $records[0].DeploymentType | Should -Be 'Collection'
        $records[0].CollectionID | Should -Be 'MCM00099'
        $records[0].PlanName | Should -Be 'Update rollout'
        $records[0].RunId | Should -Be 'RO-1'
        $records[0].Result | Should -Be 'Success'
        $records[1].DeploymentType | Should -Be 'Application'
        $records[1].DeployPurpose | Should -Be 'Required'
        $records[1].DeploymentID | Should -Be '16777300'
        $records[1].RingName | Should -Be 'Prod'
    }

    It 'uses a manual refresh when the refresh days are 0 and skips the move without a folder' {
        $s = ConvertTo-RolloutSettings -Prefs @{ RolloutRefreshDays = 0 }
        $r = Invoke-RolloutItem -Item (New-TestItem) -Settings $s -Schedule $script:Sched -RunId 'RO-2' -AuditLogPath $script:Audit
        $r.Success | Should -BeTrue
        Should -Invoke New-CMDeviceCollection -ModuleName DeploymentHelperCommon -Times 1 -ParameterFilter { $RefreshType -eq 'Manual' }
        Should -Invoke New-CMSchedule -ModuleName DeploymentHelperCommon -Times 0
        Should -Invoke Move-CMObject -ModuleName DeploymentHelperCommon -Times 0
    }

    It 'deploys to an existing collection and creates no collection' {
        $item = New-TestItem -Stage Test -Role Test -Name 'My Test PCs' -UseExisting $true
        $item.Exists = $true; $item.CollectionId = 'MCM00099'
        $r = Invoke-RolloutItem -Item $item -Settings $script:S -Schedule $script:Sched -RunId 'RO-3' -AuditLogPath $script:Audit
        $r.Success | Should -BeTrue
        $r.Created | Should -BeFalse
        Should -Invoke New-CMDeviceCollection -ModuleName DeploymentHelperCommon -Times 0
        Should -Invoke New-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 1
        @(Get-Content -LiteralPath $script:Audit).Count | Should -Be 1
    }

    It 'creates nothing and writes nothing in a dry run' {
        $r = Invoke-RolloutItem -Item (New-TestItem) -Settings $script:S -Schedule $script:Sched -RunId 'RO-4' -AuditLogPath $script:Audit -DryRun
        $r.Success | Should -BeTrue
        $r.Error | Should -Match 'Dry run'
        Should -Invoke New-CMDeviceCollection -ModuleName DeploymentHelperCommon -Times 0
        Should -Invoke New-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 0
        Test-Path -LiteralPath $script:Audit | Should -BeFalse
    }

    It 'creates nothing when the audit log cannot be written (fail closed)' {
        $dir = Join-Path $TestDrive 'a-folder'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        $r = Invoke-RolloutItem -Item (New-TestItem) -Settings $script:S -Schedule $script:Sched -RunId 'RO-5' -AuditLogPath $dir
        $r.Success | Should -BeFalse
        $r.Error | Should -Match 'Audit log unavailable'
        Should -Invoke New-CMDeviceCollection -ModuleName DeploymentHelperCommon -Times 0
    }

    It 'creates nothing when a collection of that name appeared since the check' {
        Mock Test-CollectionValid -ModuleName DeploymentHelperCommon { $script:NewCol }
        $r = Invoke-RolloutItem -Item (New-TestItem) -Settings $script:S -Schedule $script:Sched -RunId 'RO-6' -AuditLogPath $script:Audit
        $r.Success | Should -BeFalse
        $r.Error | Should -Match 'now exists'
        Should -Invoke New-CMDeviceCollection -ModuleName DeploymentHelperCommon -Times 0
    }

    It 'removes the new collection again when the query rule fails, and records the failure' {
        Mock Add-CMDeviceCollectionQueryMembershipRule -ModuleName DeploymentHelperCommon { throw [System.InvalidOperationException]::new('Invalid query expression') }
        $r = Invoke-RolloutItem -Item (New-TestItem) -Settings $script:S -Schedule $script:Sched -RunId 'RO-7' -AuditLogPath $script:Audit
        $r.Success | Should -BeFalse
        $r.Error | Should -Match 'removed again'
        $r.Error | Should -Not -Match 'Invalid query expression'
        Should -Invoke Remove-CMCollection -ModuleName DeploymentHelperCommon -Times 1 -ParameterFilter { $Id -eq 'MCM00099' }
        Should -Invoke New-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 0
        $rec = @(Get-Content -LiteralPath $script:Audit | ForEach-Object { $_ | ConvertFrom-Json })
        $rec.Count | Should -Be 1
        $rec[0].Result | Should -Match '^Failed: .*Invalid query expression'
    }

    It 'removes only the collection that it created, never another one' {
        Mock Move-CMObject -ModuleName DeploymentHelperCommon { throw 'folder' }
        Invoke-RolloutItem -Item (New-TestItem) -Settings $script:S -Schedule $script:Sched -RunId 'RO-8' -AuditLogPath $script:Audit | Out-Null
        Should -Invoke Remove-CMCollection -ModuleName DeploymentHelperCommon -Times 1 -ParameterFilter { $Id -eq 'MCM00099' }
    }

    It 'blocks the deployment when a deployment of the application appeared on the collection' {
        Mock Test-DuplicateDeployment -ModuleName DeploymentHelperCommon { [PSCustomObject]@{ AssignmentID = 5 } }
        $r = Invoke-RolloutItem -Item (New-TestItem) -Settings $script:S -Schedule $script:Sched -RunId 'RO-9' -AuditLogPath $script:Audit
        $r.Success | Should -BeFalse
        $r.Error | Should -Match 'No duplicate was created'
        Should -Invoke New-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 0
    }

    It 'refuses a built-in collection as the deployment target' {
        Mock Test-CollectionValid -ModuleName DeploymentHelperCommon { if ($CollectionId) { return [PSCustomObject]@{ Name = 'All Systems'; CollectionID = 'SMS00001'; CollectionType = 2; MemberCount = 9 } } return $null }
        $item = New-TestItem -Stage Test -Role Test -Name 'All Systems' -UseExisting $true
        $item.Exists = $true; $item.CollectionId = 'SMS00001'
        $r = Invoke-RolloutItem -Item $item -Settings $script:S -Schedule $script:Sched -RunId 'RO-10' -AuditLogPath $script:Audit
        $r.Success | Should -BeFalse
        $r.Error | Should -Match 'Built-in'
        Should -Invoke New-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 0
    }

    It 'reports a failed deployment and writes its audit record' {
        Mock New-CMApplicationDeployment -ModuleName DeploymentHelperCommon { throw 'content missing' }
        $r = Invoke-RolloutItem -Item (New-TestItem) -Settings $script:S -Schedule $script:Sched -RunId 'RO-11' -AuditLogPath $script:Audit
        $r.Success | Should -BeFalse
        $rec = @(Get-Content -LiteralPath $script:Audit | ForEach-Object { $_ | ConvertFrom-Json })
        $rec.Count | Should -Be 2
        $rec[1].Result | Should -Match '^Failed: content missing'
    }
}

Describe 'Rollout run files' {
    BeforeAll {
        function New-TestRun {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Test helper; builds test data.')]
            param([string]$State = 'TestDone', [string[]]$Statuses = @('Done', 'Planned'))
            $t = New-TestItem -Stage Test -Role Test -Name 'T'; $t.Status = $Statuses[0]
            $p = New-TestItem -Stage Prod; $p.Status = $Statuses[1]
            [PSCustomObject]@{ RunId = (New-RolloutRunId); Created = '2026-11-21T08:00:00'; State = $State; Items = @($t, $p) }
        }
    }

    It 'writes a run file and reads it back' {
        $run = New-TestRun
        $path = Save-RolloutRun -Folder (Join-Path $TestDrive 'runs') -Run $run
        $path | Should -Exist
        $back = Read-RolloutRun -Path $path
        $back.RunId | Should -Be $run.RunId
        @($back.Items).Count | Should -Be 2
        $back.Items[1].CollectionName | Should -Be 'App - Server - Installed - Version LT 2.0'
    }

    It 'refuses a file that is not a run' {
        $f = Join-Path $TestDrive 'bad.json'
        Set-Content -LiteralPath $f -Value '{"a":1}'
        { Read-RolloutRun -Path $f } | Should -Throw '*not a rollout run*'
        Set-Content -LiteralPath $f -Value 'not json'
        { Read-RolloutRun -Path $f } | Should -Throw '*could not be read*'
    }

    It 'lists open runs and leaves closed runs out unless asked' {
        $folder = Join-Path $TestDrive 'runs2'
        [void](Save-RolloutRun -Folder $folder -Run (New-TestRun -State 'TestDone'))
        [void](Save-RolloutRun -Folder $folder -Run (New-TestRun -State 'Closed'))
        @(Get-RolloutRunList -Folder $folder).Count | Should -Be 1
        @(Get-RolloutRunList -Folder $folder -IncludeClosed).Count | Should -Be 2
        @(Get-RolloutRunList -Folder (Join-Path $TestDrive 'nowhere')).Count | Should -Be 0
    }

    It 'allows a production item only after the test item of the same application and version is done' {
        @(Get-RolloutProdItem -Run (New-TestRun -Statuses 'Done', 'Planned')).Count | Should -Be 1
        @(Get-RolloutProdItem -Run (New-TestRun -Statuses 'Failed', 'Planned')).Count | Should -Be 0
        @(Get-RolloutProdItem -Run (New-TestRun -Statuses 'Planned', 'Planned')).Count | Should -Be 0
        @(Get-RolloutProdItem -Run (New-TestRun -Statuses 'Done', 'Done')).Count | Should -Be 0
    }
}


Describe 'Import-RolloutCheckResult' {
    BeforeAll {
        function Write-TestCheckCsv {
            param([string]$Path, [string[]]$Lines)
            $header = '"Application","Vendor","CMName","InstalledSoftwareName","Architecture","ConfigMgrVersion","LatestVersion","Status","Source","CheckedAtUtc"'
            Set-Content -LiteralPath $Path -Value (@($header) + $Lines) -Encoding UTF8
        }
    }

    It 'returns the rows with an update available, with the name without wildcards' {
        $f = Join-Path $TestDrive 'a.csv'
        Write-TestCheckCsv -Path $f -Lines @(
            '"7-Zip","Igor","7-Zip (x64)","7-Zip* | 7z*","x64","24.08","26.06","UpdateAvailable","APD","2026-11-20T10:00:00Z"',
            '"Audacity","","Audacity","Audacity*","x64","3.7.9","3.7.9","UpToDate","APD","2026-11-20T10:00:00Z"',
            '"Odd","","Odd","","x64","1","latest","UpdateAvailable","APD","2026-11-20T10:00:00Z"')
        $r = Import-RolloutCheckResult -Path $f
        @($r.Rows).Count | Should -Be 1
        $r.Rows[0].Application | Should -Be '7-Zip (x64)'
        $r.Rows[0].NewVersion | Should -Be '26.06'
        $r.Rows[0].SoftwareName | Should -Be '7-Zip'
        $r.Rows[0].CurrentVersion | Should -Be '24.08'
    }

    It 'takes the newer check when both apps have a row, and skips a missing file' {
        $a = Join-Path $TestDrive 'apd.csv'; $b = Join-Path $TestDrive 'pas.csv'
        Write-TestCheckCsv -Path $a -Lines @('"X","","X","X*","x64","1.0","2.0","UpdateAvailable","APD","2026-11-20T10:00:00Z"')
        Write-TestCheckCsv -Path $b -Lines @('"X","","X","X*","x64","1.0","2.5","UpdateAvailable","AppPackager","2026-11-21T10:00:00Z"')
        $r = Import-RolloutCheckResult -Path $a, $b, (Join-Path $TestDrive 'none.csv')
        @($r.Rows).Count | Should -Be 1
        $r.Rows[0].NewVersion | Should -Be '2.5'
        @((Import-RolloutCheckResult -Path (Join-Path $TestDrive 'none.csv')).Rows).Count | Should -Be 0
    }
}
