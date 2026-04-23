#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }

<#
.SYNOPSIS
    Pester 5 tests for DeploymentHelperCommon module.

.DESCRIPTION
    Unit tests with mocked ConfigurationManager cmdlets.
    Run: Invoke-Pester -Path .\Tests\DeploymentHelperCommon.Tests.ps1 -Output Detailed
#>

BeforeAll {
    # Stub CM cmdlets so Pester can mock them without the CM console installed.
    # These do nothing by default -- each test overrides them via Mock.
    $cmStubs = @(
        'Get-CMSite', 'Get-CMApplication', 'Get-CMCollection',
        'Get-CMSoftwareUpdateGroup', 'Get-CMDistributionStatus',
        'Get-CMDistributionPointGroup', 'Get-CMDeployment',
        'Get-CMApplicationDeployment',
        'Start-CMContentDistribution',
        'New-CMApplicationDeployment', 'New-CMSoftwareUpdateDeployment'
    )
    foreach ($name in $cmStubs) {
        if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
            Set-Item -Path "function:global:$name" -Value { }
        }
    }

    $modulePath = Join-Path $PSScriptRoot '..\Module\DeploymentHelperCommon.psd1'
    Import-Module $modulePath -Force

    # Temp paths for test artifacts
    $script:TestLogPath = Join-Path $TestDrive 'test.log'
    $script:TestDeployLogPath = Join-Path $TestDrive 'deployment-log.jsonl'
    $script:TestTemplatePath = Join-Path $TestDrive 'Templates'
    New-Item -ItemType Directory -Path $script:TestTemplatePath -Force | Out-Null
}

Describe 'Initialize-Logging' {
    It 'Creates log file with header' {
        Initialize-Logging -LogPath $script:TestLogPath
        $script:TestLogPath | Should -Exist
        $content = Get-Content -LiteralPath $script:TestLogPath -Raw
        $content | Should -Match 'Log initialized'
    }

    It 'Creates parent directory if missing' {
        $nestedPath = Join-Path $TestDrive 'sub\dir\test.log'
        Initialize-Logging -LogPath $nestedPath
        $nestedPath | Should -Exist
    }
}

Describe 'Write-Log' {
    BeforeAll {
        Initialize-Logging -LogPath $script:TestLogPath
    }

    It 'Writes INFO message to log file' {
        Write-Log 'Test info message' -Quiet
        $content = Get-Content -LiteralPath $script:TestLogPath -Raw
        $content | Should -Match 'INFO.*Test info message'
    }

    It 'Writes WARN message to log file' {
        Write-Log 'Test warning' -Level WARN -Quiet
        $content = Get-Content -LiteralPath $script:TestLogPath -Raw
        $content | Should -Match 'WARN.*Test warning'
    }

    It 'Writes ERROR message to log file' {
        Write-Log 'Test error' -Level ERROR -Quiet
        $content = Get-Content -LiteralPath $script:TestLogPath -Raw
        $content | Should -Match 'ERROR.*Test error'
    }

    It 'Accepts empty string message' {
        { Write-Log '' -Quiet } | Should -Not -Throw
    }
}

Describe 'Connect-CMSite' {
    BeforeAll {
        Mock Import-Module { } -ModuleName DeploymentHelperCommon -ParameterFilter { $Name -like '*ConfigurationManager*' }
        Mock Get-Module { $null } -ModuleName DeploymentHelperCommon -ParameterFilter { $Name -eq 'ConfigurationManager' }
        Mock Get-PSDrive { $null } -ModuleName DeploymentHelperCommon -ParameterFilter { $PSProvider -eq 'CMSite' }
        Mock New-PSDrive { [PSCustomObject]@{ Name = 'MCM' } } -ModuleName DeploymentHelperCommon
        Mock Set-Location { } -ModuleName DeploymentHelperCommon
        Mock Get-CMSite { [PSCustomObject]@{ SiteCode = 'MCM'; SiteName = 'Test Site' } } -ModuleName DeploymentHelperCommon
    }

    It 'Returns true on successful connection' {
        $env:SMS_ADMIN_UI_PATH = $TestDrive
        Set-Content -Path (Join-Path $TestDrive 'ConfigurationManager.psd1') -Value '@{ ModuleVersion = "1.0" }'

        $result = Connect-CMSite -SiteCode 'MCM' -SMSProvider 'sms.example.com'
        $result | Should -BeTrue
    }

    It 'Returns false when CM module not found' {
        $savedPath = $env:SMS_ADMIN_UI_PATH
        $env:SMS_ADMIN_UI_PATH = $null
        $result = Connect-CMSite -SiteCode 'MCM' -SMSProvider 'sms.example.com'
        $result | Should -BeFalse
        $env:SMS_ADMIN_UI_PATH = $savedPath
    }
}

Describe 'Test-CMConnection' {
    It 'Returns false when not connected' {
        Disconnect-CMSite
        $result = Test-CMConnection
        $result | Should -BeFalse
    }
}

Describe 'Test-ApplicationExists' {
    Context 'Application found' {
        BeforeAll {
            Mock Get-CMApplication -ModuleName DeploymentHelperCommon {
                [PSCustomObject]@{
                    LocalizedDisplayName = '7-Zip'
                    SoftwareVersion      = '24.09'
                    PackageID            = 'MCM00001'
                }
            }
        }

        It 'Returns application object when found' {
            $result = Test-ApplicationExists -ApplicationName '7-Zip'
            $result | Should -Not -BeNullOrEmpty
            $result.LocalizedDisplayName | Should -Be '7-Zip'
            $result.SoftwareVersion | Should -Be '24.09'
        }
    }

    Context 'Application not found' {
        BeforeAll {
            Mock Get-CMApplication -ModuleName DeploymentHelperCommon { $null }
        }

        It 'Returns null when not found' {
            $result = Test-ApplicationExists -ApplicationName 'NonExistent'
            $result | Should -BeNullOrEmpty
        }
    }

    Context 'CM cmdlet throws' {
        BeforeAll {
            Mock Get-CMApplication -ModuleName DeploymentHelperCommon { throw 'Connection error' }
        }

        It 'Returns null on error' {
            $result = Test-ApplicationExists -ApplicationName 'Broken'
            $result | Should -BeNullOrEmpty
        }
    }
}

Describe 'Search-CMApplicationByName' {
    Context 'Results found' {
        BeforeAll {
            Mock Get-CMApplication -ModuleName DeploymentHelperCommon {
                @(
                    [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; SoftwareVersion = '24.09'; PackageID = 'MCM00001'; DateLastModified = (Get-Date) },
                    [PSCustomObject]@{ LocalizedDisplayName = '7-Zip (x86)'; SoftwareVersion = '24.09'; PackageID = 'MCM00002'; DateLastModified = (Get-Date) }
                )
            }
        }

        It 'Returns matching applications' {
            $results = Search-CMApplicationByName -SearchText '7-Zip'
            @($results).Count | Should -Be 2
        }
    }

    Context 'No results' {
        BeforeAll {
            Mock Get-CMApplication -ModuleName DeploymentHelperCommon { @() }
        }

        It 'Returns empty array' {
            $results = Search-CMApplicationByName -SearchText 'XYZZY'
            @($results).Count | Should -Be 0
        }
    }
}

Describe 'Search-CMCollectionByName' {
    Context 'Results found' {
        BeforeAll {
            Mock Get-CMCollection -ModuleName DeploymentHelperCommon {
                @(
                    [PSCustomObject]@{ Name = 'Pilot Workstations'; CollectionID = 'MCM00010'; MemberCount = 5; LastRefreshTime = (Get-Date) },
                    [PSCustomObject]@{ Name = 'Pilot Servers'; CollectionID = 'MCM00011'; MemberCount = 2; LastRefreshTime = (Get-Date) }
                )
            }
        }

        It 'Returns matching device collections' {
            $results = Search-CMCollectionByName -SearchText 'Pilot'
            @($results).Count | Should -Be 2
        }
    }
}

Describe 'Test-CollectionValid' {
    Context 'Valid device collection' {
        BeforeAll {
            Mock Get-CMCollection -ModuleName DeploymentHelperCommon {
                [PSCustomObject]@{ Name = 'Test Collection'; CollectionID = 'MCM00010'; CollectionType = 2; MemberCount = 3 }
            }
        }

        It 'Returns collection object for device collection' {
            $result = Test-CollectionValid -CollectionName 'Test Collection'
            $result | Should -Not -BeNullOrEmpty
            $result.CollectionType | Should -Be 2
        }
    }

    Context 'User collection rejected' {
        BeforeAll {
            Mock Get-CMCollection -ModuleName DeploymentHelperCommon {
                [PSCustomObject]@{ Name = 'User Collection'; CollectionID = 'MCM00020'; CollectionType = 1; MemberCount = 10 }
            }
        }

        It 'Returns null for user collection' {
            $result = Test-CollectionValid -CollectionName 'User Collection'
            $result | Should -BeNullOrEmpty
        }
    }

    Context 'Collection not found' {
        BeforeAll {
            Mock Get-CMCollection -ModuleName DeploymentHelperCommon { $null }
        }

        It 'Returns null when not found' {
            $result = Test-CollectionValid -CollectionName 'Ghost'
            $result | Should -BeNullOrEmpty
        }
    }
}

Describe 'Test-CollectionSafe' {
    It 'Blocks SMS00001 (All Systems)' {
        $col = [PSCustomObject]@{ Name = 'All Systems'; CollectionID = 'SMS00001' }
        $result = Test-CollectionSafe -Collection $col
        $result.IsSafe | Should -BeFalse
        $result.Reason | Should -Match 'system collection'
    }

    It 'Blocks SMS00004 (All Unknown Computers)' {
        $col = [PSCustomObject]@{ Name = 'All Unknown Computers'; CollectionID = 'SMS00004' }
        $result = Test-CollectionSafe -Collection $col
        $result.IsSafe | Should -BeFalse
    }

    It 'Blocks any SMS000* pattern' {
        $col = [PSCustomObject]@{ Name = 'All Desktop and Server Clients'; CollectionID = 'SMS000C1' }
        $result = Test-CollectionSafe -Collection $col
        $result.IsSafe | Should -BeFalse
    }

    It 'Allows custom collections' {
        $col = [PSCustomObject]@{ Name = 'Pilot Workstations'; CollectionID = 'MCM00010' }
        $result = Test-CollectionSafe -Collection $col
        $result.IsSafe | Should -BeTrue
    }

    It 'Allows SMSDM-prefixed collections' {
        $col = [PSCustomObject]@{ Name = 'All Desktop and Server Clients'; CollectionID = 'SMSDM003' }
        $result = Test-CollectionSafe -Collection $col
        $result.IsSafe | Should -BeTrue
    }
}

Describe 'Test-DuplicateDeployment' {
    Context 'No duplicate exists' {
        BeforeAll {
            Mock Get-CMApplicationDeployment -ModuleName DeploymentHelperCommon { $null }
        }

        It 'Returns null when no duplicate' {
            $result = Test-DuplicateDeployment -ApplicationName '7-Zip' -CollectionName 'Pilot'
            $result | Should -BeNullOrEmpty
        }
    }

    Context 'Duplicate exists' {
        BeforeAll {
            Mock Get-CMApplicationDeployment -ModuleName DeploymentHelperCommon {
                [PSCustomObject]@{ AssignmentID = 12345; ApplicationName = '7-Zip'; CollectionName = 'Pilot' }
            }
        }

        It 'Returns deployment object when duplicate found' {
            $result = Test-DuplicateDeployment -ApplicationName '7-Zip' -CollectionName 'Pilot'
            $result | Should -Not -BeNullOrEmpty
        }
    }
}

Describe 'Test-ContentDistributed' {
    Context 'Fully distributed' {
        BeforeAll {
            Mock Get-CMDistributionStatus -ModuleName DeploymentHelperCommon {
                [PSCustomObject]@{ Targeted = 3; NumberSuccess = 3; NumberInProgress = 0; NumberErrors = 0 }
            }
        }

        It 'Returns IsFullyDistributed true when all DPs succeeded' {
            $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; PackageID = 'MCM00001' }
            $result = Test-ContentDistributed -Application $app
            $result.IsFullyDistributed | Should -BeTrue
            $result.NumberSuccess | Should -Be 3
            $result.Targeted | Should -Be 3
        }
    }

    Context 'Partially distributed' {
        BeforeAll {
            Mock Get-CMDistributionStatus -ModuleName DeploymentHelperCommon {
                [PSCustomObject]@{ Targeted = 3; NumberSuccess = 2; NumberInProgress = 0; NumberErrors = 1 }
            }
        }

        It 'Returns IsFullyDistributed false when errors exist' {
            $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; PackageID = 'MCM00001' }
            $result = Test-ContentDistributed -Application $app
            $result.IsFullyDistributed | Should -BeFalse
            $result.NumberErrors | Should -Be 1
        }
    }

    Context 'No distribution status' {
        BeforeAll {
            Mock Get-CMDistributionStatus -ModuleName DeploymentHelperCommon { $null }
        }

        It 'Returns not distributed when no status exists' {
            $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; PackageID = 'MCM00001' }
            $result = Test-ContentDistributed -Application $app
            $result.IsFullyDistributed | Should -BeFalse
            $result.Targeted | Should -Be 0
        }
    }
}

Describe 'Get-DPGroupList' {
    BeforeAll {
        Mock Get-CMDistributionPointGroup -ModuleName DeploymentHelperCommon {
            @(
                [PSCustomObject]@{ Name = 'All Distribution Points'; GroupID = 1 },
                [PSCustomObject]@{ Name = 'Branch Office DPs'; GroupID = 2 }
            )
        }
    }

    It 'Returns DP groups' {
        $result = Get-DPGroupList
        @($result).Count | Should -Be 2
    }
}

Describe 'Start-ContentDistributionToGroups' {
    Context 'Successful distribution' {
        BeforeAll {
            Mock Start-CMContentDistribution -ModuleName DeploymentHelperCommon { }
        }

        It 'Returns success for each group' {
            $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip' }
            $results = Start-ContentDistributionToGroups -Application $app -DPGroupNames @('Group1', 'Group2')
            @($results).Count | Should -Be 2
            $results[0].Success | Should -BeTrue
            $results[1].Success | Should -BeTrue
        }
    }

    Context 'Already targeted' {
        BeforeAll {
            Mock Start-CMContentDistribution -ModuleName DeploymentHelperCommon { throw 'already been targeted' }
        }

        It 'Handles already-targeted gracefully' {
            $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip' }
            $results = @(Start-ContentDistributionToGroups -Application $app -DPGroupNames @('Group1'))
            $results[0].Success | Should -BeTrue
            $results[0].AlreadyTargeted | Should -BeTrue
        }
    }

    Context 'Distribution error' {
        BeforeAll {
            Mock Start-CMContentDistribution -ModuleName DeploymentHelperCommon { throw 'Network error' }
        }

        It 'Returns failure with error message' {
            $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip' }
            $results = @(Start-ContentDistributionToGroups -Application $app -DPGroupNames @('Group1'))
            $results[0].Success | Should -BeFalse
            $results[0].Error | Should -Match 'Network error'
        }
    }
}

Describe 'Test-SUGExists' {
    Context 'SUG found' {
        BeforeAll {
            Mock Get-CMSoftwareUpdateGroup -ModuleName DeploymentHelperCommon {
                [PSCustomObject]@{ LocalizedDisplayName = '2026-04 Updates'; NumberOfUpdates = 15; NumberOfExpiredUpdates = 0 }
            }
        }

        It 'Returns SUG object' {
            $result = Test-SUGExists -SUGName '2026-04 Updates'
            $result | Should -Not -BeNullOrEmpty
            $result.NumberOfUpdates | Should -Be 15
        }
    }

    Context 'SUG not found' {
        BeforeAll {
            Mock Get-CMSoftwareUpdateGroup -ModuleName DeploymentHelperCommon { $null }
        }

        It 'Returns null' {
            $result = Test-SUGExists -SUGName 'Nonexistent'
            $result | Should -BeNullOrEmpty
        }
    }
}

Describe 'Get-DeploymentPreview' {
    It 'Returns application preview' {
        $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; SoftwareVersion = '24.09' }
        $col = [PSCustomObject]@{ Name = 'Pilot'; CollectionID = 'MCM00010'; MemberCount = 5 }
        $result = Get-DeploymentPreview -TargetObject $app -Collection $col -DeploymentType 'Application'
        $result.ApplicationName | Should -Be '7-Zip'
        $result.ApplicationVersion | Should -Be '24.09'
        $result.CollectionName | Should -Be 'Pilot'
        $result.MemberCount | Should -Be 5
    }

    It 'Returns SUG preview with update count' {
        $sug = [PSCustomObject]@{ LocalizedDisplayName = '2026-04 Updates'; NumberOfUpdates = 15 }
        $col = [PSCustomObject]@{ Name = 'All Workstations'; CollectionID = 'MCM00020'; MemberCount = 100 }
        $result = Get-DeploymentPreview -TargetObject $sug -Collection $col -DeploymentType 'SUG'
        $result.ApplicationVersion | Should -Match '15 updates'
    }
}

Describe 'Invoke-ApplicationDeployment' {
    Context 'Successful deployment' {
        BeforeAll {
            Mock New-CMApplicationDeployment -ModuleName DeploymentHelperCommon {
                [PSCustomObject]@{ AssignmentID = 16777300; AssignmentUniqueID = '{GUID}' }
            }
        }

        It 'Returns success with deployment ID' {
            $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; SoftwareVersion = '24.09' }
            $col = [PSCustomObject]@{ Name = 'Pilot'; MemberCount = 5 }
            $result = Invoke-ApplicationDeployment -Application $app -Collection $col `
                -DeployPurpose 'Available' -AvailableDateTime (Get-Date)
            $result.Success | Should -BeTrue
            $result.DeploymentID | Should -Be 16777300
        }

        It 'Calls New-CMApplicationDeployment for Available' {
            $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; SoftwareVersion = '24.09' }
            $col = [PSCustomObject]@{ Name = 'Pilot'; MemberCount = 5 }
            Invoke-ApplicationDeployment -Application $app -Collection $col `
                -DeployPurpose 'Available' -AvailableDateTime (Get-Date) `
                -TimeBasedOn 'Utc' -UserNotification 'DisplaySoftwareCenterOnly'

            Should -Invoke New-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 1
        }

        It 'Calls New-CMApplicationDeployment for Required with deadline' {
            $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; SoftwareVersion = '24.09' }
            $col = [PSCustomObject]@{ Name = 'Pilot'; MemberCount = 5 }
            $result = Invoke-ApplicationDeployment -Application $app -Collection $col `
                -DeployPurpose 'Required' -AvailableDateTime (Get-Date) `
                -DeadlineDateTime (Get-Date).AddHours(24) -UseMeteredNetwork $true

            $result.Success | Should -BeTrue
            Should -Invoke New-CMApplicationDeployment -ModuleName DeploymentHelperCommon -Times 1
        }
    }

    Context 'Failed deployment' {
        BeforeAll {
            Mock New-CMApplicationDeployment -ModuleName DeploymentHelperCommon { throw 'Deployment creation failed' }
        }

        It 'Returns failure with error message' {
            $app = [PSCustomObject]@{ LocalizedDisplayName = '7-Zip'; SoftwareVersion = '24.09' }
            $col = [PSCustomObject]@{ Name = 'Pilot'; MemberCount = 5 }
            $result = Invoke-ApplicationDeployment -Application $app -Collection $col `
                -DeployPurpose 'Available' -AvailableDateTime (Get-Date)
            $result.Success | Should -BeFalse
            $result.Error | Should -Match 'Deployment creation failed'
        }
    }
}

Describe 'Invoke-SUGDeployment' {
    Context 'Successful SUG deployment' {
        BeforeAll {
            Mock New-CMSoftwareUpdateDeployment -ModuleName DeploymentHelperCommon {
                [PSCustomObject]@{ AssignmentID = 16777400 }
            }
        }

        It 'Returns success with deployment ID' {
            $sug = [PSCustomObject]@{ LocalizedDisplayName = '2026-04 Updates'; NumberOfUpdates = 15 }
            $col = [PSCustomObject]@{ Name = 'All Workstations'; MemberCount = 100 }
            $result = Invoke-SUGDeployment -SUG $sug -Collection $col `
                -DeployPurpose 'Available' -AvailableDateTime (Get-Date)
            $result.Success | Should -BeTrue
            $result.DeploymentID | Should -Be 16777400
        }

        It 'Calls New-CMSoftwareUpdateDeployment for Required with options' {
            $sug = [PSCustomObject]@{ LocalizedDisplayName = '2026-04 Updates'; NumberOfUpdates = 15 }
            $col = [PSCustomObject]@{ Name = 'All Workstations'; MemberCount = 100 }
            $result = Invoke-SUGDeployment -SUG $sug -Collection $col `
                -DeployPurpose 'Required' -AvailableDateTime (Get-Date) `
                -DeadlineDateTime (Get-Date).AddDays(7) `
                -DownloadFromMicrosoftUpdate $true -AllowBoundaryFallback $true

            $result.Success | Should -BeTrue
            Should -Invoke New-CMSoftwareUpdateDeployment -ModuleName DeploymentHelperCommon -Times 1
        }
    }
}

Describe 'Write-DeploymentLog' {
    It 'Creates JSONL file with correct fields' {
        $logPath = Join-Path $TestDrive 'deploy-test.jsonl'
        Write-DeploymentLog -LogPath $logPath -Record @{
            DeploymentType     = 'Application'
            ApplicationName    = '7-Zip'
            ApplicationVersion = '24.09'
            CollectionName     = 'Pilot'
            CollectionID       = 'MCM00010'
            MemberCount        = 5
            DeployPurpose      = 'Available'
            DeadlineDateTime   = ''
            DeploymentID       = '16777300'
            Result             = 'Success'
        }
        $logPath | Should -Exist
        $content = Get-Content -LiteralPath $logPath -Raw
        $entry = $content.Trim() | ConvertFrom-Json
        $entry.DeploymentType | Should -Be 'Application'
        $entry.ApplicationName | Should -Be '7-Zip'
        $entry.DeployAction | Should -Be 'Install'
        $entry.User | Should -Not -BeNullOrEmpty
        $entry.Timestamp | Should -Match '^\d{4}-\d{2}-\d{2}T'
    }

    It 'Appends to existing log' {
        $logPath = Join-Path $TestDrive 'deploy-append.jsonl'
        $record = @{
            DeploymentType = 'Application'; ApplicationName = 'App1'; ApplicationVersion = '1.0'
            CollectionName = 'Col1'; CollectionID = 'C1'; MemberCount = 1
            DeployPurpose = 'Available'; DeadlineDateTime = ''; DeploymentID = '1'; Result = 'Success'
        }
        Write-DeploymentLog -LogPath $logPath -Record $record
        $record.ApplicationName = 'App2'
        $record.DeploymentID = '2'
        Write-DeploymentLog -LogPath $logPath -Record $record

        $lines = Get-Content -LiteralPath $logPath
        $lines.Count | Should -Be 2
    }
}

Describe 'Get-DeploymentHistory' {
    It 'Reads JSONL records' {
        $logPath = Join-Path $TestDrive 'history-test.jsonl'
        $record = @{
            DeploymentType = 'Application'; ApplicationName = 'TestApp'; ApplicationVersion = '1.0'
            CollectionName = 'TestCol'; CollectionID = 'T1'; MemberCount = 3
            DeployPurpose = 'Required'; DeadlineDateTime = '2026-04-15T00:00:00'; DeploymentID = '99'
            Result = 'Success'
        }
        Write-DeploymentLog -LogPath $logPath -Record $record
        Write-DeploymentLog -LogPath $logPath -Record $record

        $history = Get-DeploymentHistory -LogPath $logPath
        $history.Count | Should -Be 2
        $history[0].ApplicationName | Should -Be 'TestApp'
    }

    It 'Returns empty array for missing log file' {
        $history = Get-DeploymentHistory -LogPath (Join-Path $TestDrive 'nonexistent.jsonl')
        $history.Count | Should -Be 0
    }
}

Describe 'Get-DeploymentTemplates' {
    It 'Loads JSON template files' {
        $tmplDir = Join-Path $TestDrive 'tmpl-test'
        New-Item -ItemType Directory -Path $tmplDir -Force | Out-Null
        @{
            Name = 'Test Template'
            DeployPurpose = 'Available'
            UserNotification = 'DisplayAll'
            TimeBasedOn = 'LocalTime'
            OverrideServiceWindow = $false
            RebootOutsideServiceWindow = $false
            DefaultDeadlineOffsetHours = 0
        } | ConvertTo-Json | Set-Content (Join-Path $tmplDir 'Test.json') -Encoding UTF8

        $templates = Get-DeploymentTemplates -TemplatePath $tmplDir
        @($templates).Count | Should -Be 1
        $templates[0].Name | Should -Be 'Test Template'
    }

    It 'Returns empty for missing directory' {
        $templates = Get-DeploymentTemplates -TemplatePath (Join-Path $TestDrive 'no-such-dir')
        @($templates).Count | Should -Be 0
    }
}

Describe 'Save-DeploymentTemplate' {
    It 'Saves template with all fields' {
        $tmplPath = Join-Path $TestDrive 'saved-template.json'
        Save-DeploymentTemplate -TemplatePath $tmplPath -TemplateName 'My Template' -Config @{
            DeployPurpose              = 'Required'
            UserNotification           = 'HideAll'
            TimeBasedOn                = 'Utc'
            OverrideServiceWindow      = $true
            RebootOutsideServiceWindow = $true
            AllowMeteredConnection     = $true
            AllowBoundaryFallback      = $true
            AllowMicrosoftUpdate       = $false
            RequirePostRebootFullScan  = $true
            DefaultDeadlineOffsetHours = 168
        }

        $tmplPath | Should -Exist
        $loaded = Get-Content -LiteralPath $tmplPath -Raw | ConvertFrom-Json
        $loaded.Name | Should -Be 'My Template'
        $loaded.DeployPurpose | Should -Be 'Required'
        $loaded.TimeBasedOn | Should -Be 'Utc'
        $loaded.DefaultDeadlineOffsetHours | Should -Be 168
    }
}

Describe 'Export-DeploymentHistoryCsv' {
    It 'Exports records to CSV' {
        $csvPath = Join-Path $TestDrive 'export.csv'
        $records = @(
            [PSCustomObject]@{ Timestamp = '2026-04-14'; ApplicationName = '7-Zip'; Result = 'Success' },
            [PSCustomObject]@{ Timestamp = '2026-04-14'; ApplicationName = 'Notepad++'; Result = 'Failed' }
        )
        Export-DeploymentHistoryCsv -Records $records -OutputPath $csvPath
        $csvPath | Should -Exist
        $imported = Import-Csv -LiteralPath $csvPath
        $imported.Count | Should -Be 2
    }
}

Describe 'Export-DeploymentHistoryHtml' {
    It 'Exports records to HTML with styling' {
        $htmlPath = Join-Path $TestDrive 'export.html'
        $records = @(
            [PSCustomObject]@{ Timestamp = '2026-04-14'; User = 'TEST\Admin'; DeploymentType = 'Application'; ApplicationName = '7-Zip'; ApplicationVersion = '24.09'; CollectionName = 'Pilot'; MemberCount = 5; DeployPurpose = 'Available'; DeadlineDateTime = ''; DeploymentID = '1'; Result = 'Success' }
        )
        Export-DeploymentHistoryHtml -Records $records -OutputPath $htmlPath
        $htmlPath | Should -Exist
        $content = Get-Content -LiteralPath $htmlPath -Raw
        $content | Should -Match '<table>'
        $content | Should -Match 'success'
        $content | Should -Match 'DeploymentType'
    }
}
