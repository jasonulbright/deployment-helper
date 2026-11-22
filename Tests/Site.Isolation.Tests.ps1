#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }

<#
.SYNOPSIS
    Two-site isolation of the per-site data of Deployment Helper (Q-112f).

.DESCRIPTION
    Each test seeds an old-layout tool folder in TestDrive (prefs, audit log, ring plans, ring runs, rollout runs),
    points APP_PACKAGER_DATA_ROOT at TestDrive, and starts the tool logic for site AB1, then CD2, then AB1 again.
    The tool logic is the real code: Resolve-SuiteToolSite from the vendored SuiteCommon, Get-DhPreferences and
    Save-DhPreferences of the start script (cut out of its syntax tree, as in Shell.Logic.Tests.ps1), and
    Resolve-DhStorePath of the module. Site codes are fictional (AB1, CD2). Hosts are example.com.
    No window opens and no ConfigMgr connection is needed.
#>

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:ShellPath = Join-Path $script:AppRoot 'start-deploymenthelper.ps1'
    $script:ShellText = Get-Content -LiteralPath $script:ShellPath -Raw
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
    foreach ($n in 'Get-DhPreferences', 'Save-DhPreferences') {
        $def = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true)
        if (-not $def) { throw "function $n not found in the script" }
        # The functions read $PSScriptRoot, which is empty in a scriptblock that was cut out. It points to the seeded tool folder here.
        . ([scriptblock]::Create($def.Extent.Text.Replace('$PSScriptRoot', '$script:ToolRoot')))
    }
    $script:EnvNames = 'SUITE_CM_SITECODE', 'APP_PACKAGER_SITE', 'APP_PACKAGER_DATA_ROOT', 'SUITE_CM_PROVIDER'

    function Write-Seed {
        param([string]$Root, [string]$Relative, [string]$Text)
        $p = Join-Path $Root $Relative
        $d = Split-Path $p -Parent
        if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
        [IO.File]::WriteAllText($p, $Text, (New-Object Text.UTF8Encoding($false)))
    }
    function Get-TreeHash {
        param([string]$Root)
        $h = @{}
        if (Test-Path -LiteralPath $Root) { foreach ($f in Get-ChildItem -LiteralPath $Root -Recurse -File -Force) { $h[$f.FullName] = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash } }
        return $h
    }
    # An old tool folder beside the script, as the tool wrote it before the per-site layout.
    function New-OldToolFolder {
        param([string]$Suite, [switch]$NoSiteCode)
        $t = Join-Path $Suite 'deployment-helper'
        $prefs = [ordered]@{
            RolloutNameTemplate    = '{App} - {Role} - old'
            DeploymentAuditLogPath = (Join-Path $t 'Logs\deployment-audit.jsonl')
            RingRunStatePath       = (Join-Path $t 'Rings\runs')
            RolloutRunPath         = (Join-Path $t 'Rollouts')
        }
        if (-not $NoSiteCode) { $prefs['SiteCode'] = 'AB1'; $prefs['SMSProvider'] = 'cm01.example.com' }
        Write-Seed $t 'DeploymentHelper.prefs.json' ($prefs | ConvertTo-Json)
        Write-Seed $t 'Logs\deployment-audit.jsonl' ('{"Timestamp":"2026-11-20T10:00:00","ApplicationName":"App AB1","Result":"Success"}' + "`n")
        Write-Seed $t 'Logs\DeploymentHelper-20261120-100000.log' 'global log'
        Write-Seed $t 'Rings\AB1-plan.json' '{"Name":"AB1-plan","Rings":[]}'
        Write-Seed $t 'Rings\runs\run-ab1.json' '{"RunId":"run-ab1"}'
        Write-Seed $t 'Rollouts\Rollout-ab1.json' '{"RunId":"ab1"}'
        Write-Seed $t 'Templates\Pilot.json' '{"Name":"Pilot"}'
        return $t
    }
    # What the start script does at start: resolve the site, read the prefs, resolve the four stores.
    function Start-Tool {
        param([string]$Folder)
        if ($Folder) { [Environment]::SetEnvironmentVariable('APP_PACKAGER_SITE', $Folder) } else { [Environment]::SetEnvironmentVariable('APP_PACKAGER_SITE', $null) }
        $global:PrefsPath = Join-Path $script:ToolRoot 'DeploymentHelper.prefs.json'
        $script:Site = Resolve-SuiteToolSite -Tool 'deployment-helper' -ToolRoot $script:ToolRoot -LegacyPrefsPath $global:PrefsPath
        $global:Prefs = Get-DhPreferences
        $r = @{ Site = $script:Site; Prefs = $global:Prefs }
        $r.AuditLog = Resolve-DhStorePath -Store AuditLog -ToolRoot $script:ToolRoot -Pref $global:Prefs['DeploymentAuditLogPath'] -Site $script:Site -Create
        $r.RingPlans = Resolve-DhStorePath -Store RingPlans -ToolRoot $script:ToolRoot -Site $script:Site -Create
        $r.RingRuns = Resolve-DhStorePath -Store RingRuns -ToolRoot $script:ToolRoot -Pref $global:Prefs['RingRunStatePath'] -Site $script:Site
        $r.Rollouts = Resolve-DhStorePath -Store Rollouts -ToolRoot $script:ToolRoot -Pref $global:Prefs['RolloutRunPath'] -Site $script:Site
        return $r
    }
}

AfterAll {
    Remove-Variable -Name PrefsPath, Prefs -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Per-site data of Deployment Helper' {
    BeforeAll {
        $script:Saved = @{}
        foreach ($n in $script:EnvNames) { $script:Saved[$n] = [Environment]::GetEnvironmentVariable($n); [Environment]::SetEnvironmentVariable($n, $null) }
        $temp = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $script:Suite = Join-Path $temp 'AppPackagerSuite'
        $script:Data = Join-Path $temp 'Data'
        [void](New-Item -ItemType Directory -Path $script:Suite, $script:Data -Force)
        [Environment]::SetEnvironmentVariable('APP_PACKAGER_DATA_ROOT', $script:Data)
        $script:ToolRoot = New-OldToolFolder -Suite $script:Suite
        $script:OldHash = Get-TreeHash -Root $script:ToolRoot
        $script:DataRoot = Get-SuiteDataRoot

        # Start for AB1, use the tool, save a changed provider and write one audit line.
        $script:Ab1 = Start-Tool -Folder 'AB1'
        $script:Ab1Prefs = $script:Ab1.Prefs.Clone()
        $global:Prefs['SMSProvider'] = 'cm02.example.com'
        Save-DhPreferences -Prefs $global:Prefs
        [void](Write-DeploymentLog -LogPath $script:Ab1.AuditLog -Record @{ DeploymentType = 'Application'; ApplicationName = 'New AB1 app'; Result = 'Success' })
        Write-Seed $script:Ab1.RingRuns 'run-ab1-new.json' '{"RunId":"run-ab1-new"}'
        $script:Ab1After = Get-TreeHash -Root (Join-Path $script:DataRoot 'sites\AB1')

        # Start for CD2 and write there.
        $script:Cd2 = Start-Tool -Folder 'CD2'
        $script:Cd2Prefs = $script:Cd2.Prefs.Clone()
        [void](Write-DeploymentLog -LogPath $script:Cd2.AuditLog -Record @{ DeploymentType = 'Application'; ApplicationName = 'CD2 app'; Result = 'Success' })
        Write-Seed $script:Cd2.RingRuns 'run-cd2.json' '{"RunId":"run-cd2"}'
        $global:Prefs['SMSProvider'] = 'cm03.example.com'
        Save-DhPreferences -Prefs $global:Prefs

        # Back to AB1.
        $script:Ab1Again = Start-Tool -Folder 'AB1'
    }

    AfterAll {
        foreach ($n in $script:EnvNames) { [Environment]::SetEnvironmentVariable($n, $script:Saved[$n]) }
    }

    It 'keeps every old file unchanged (SHA-256), except the prefs file, which only gets global keys' {
        $now = Get-TreeHash -Root $script:ToolRoot
        foreach ($k in $script:OldHash.Keys) {
            $now.ContainsKey($k) | Should -BeTrue -Because "$k must still exist"
            if ($k -like '*DeploymentHelper.prefs.json') { continue }
            $now[$k] | Should -Be $script:OldHash[$k] -Because "$k must keep its content"
        }
        $oldPrefs = Get-Content -LiteralPath (Join-Path $script:ToolRoot 'DeploymentHelper.prefs.json') -Raw | ConvertFrom-Json
        $oldPrefs.SMSProvider | Should -Be 'cm01.example.com' -Because 'the old file keeps the site keys it already had'
        $oldPrefs.RolloutNameTemplate | Should -Be '{App} - {Role} - old'
    }

    It 'copies the AB1 stores once into sites\AB1\tools\deployment-helper' {
        $tool = Join-Path $script:DataRoot 'sites\AB1\tools\deployment-helper'
        $script:Ab1.Site.Active | Should -BeTrue
        $script:Ab1.Site.ToolFolder | Should -Be $tool
        (Get-Content -LiteralPath (Join-Path $tool 'deployment-audit.jsonl') -Raw) | Should -Match 'App AB1'
        Test-Path -LiteralPath (Join-Path $tool 'rings\AB1-plan.json') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $tool 'rings\runs\run-ab1.json') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $tool 'rollouts\Rollout-ab1.json') | Should -BeTrue
        $script:Ab1.AuditLog | Should -Be (Join-Path $tool 'deployment-audit.jsonl')
        $script:Ab1.RingPlans | Should -Be (Join-Path $tool 'rings')
        $script:Ab1.RingRuns | Should -Be (Join-Path $tool 'rings\runs')
        $script:Ab1.Rollouts | Should -Be (Join-Path $tool 'rollouts')
        (Get-Content -LiteralPath (Join-Path $tool 'migrated.json') -Raw | ConvertFrom-Json).tool | Should -Be 'deployment-helper'
    }

    It 'reads the AB1 site keys from the old prefs and stores no absolute default path' {
        $script:Ab1Prefs['SiteCode'] | Should -Be 'AB1'
        $script:Ab1Prefs['SMSProvider'] | Should -Be 'cm01.example.com'
        $script:Ab1Prefs['DeploymentAuditLogPath'] | Should -Be ''
        $script:Ab1Prefs['RingRunStatePath'] | Should -Be ''
        $script:Ab1Prefs['RolloutRunPath'] | Should -Be ''
        $script:Ab1Prefs['RolloutNameTemplate'] | Should -Be '{App} - {Role} - old'
        $siteFile = Get-Content -LiteralPath (Join-Path $script:DataRoot 'sites\AB1\tools\deployment-helper\prefs.json') -Raw | ConvertFrom-Json
        $siteFile.SiteCode | Should -Be 'AB1'
        $siteFile.PSObject.Properties.Name | Should -Not -Contain 'DeploymentAuditLogPath'
        $siteFile.PSObject.Properties.Name | Should -Not -Contain 'RingRunStatePath'
        $siteFile.PSObject.Properties.Name | Should -Not -Contain 'RolloutRunPath'
    }

    It 'shows none of the AB1 values or files for CD2' {
        $script:Cd2.Site.SiteFolder | Should -Be 'CD2'
        $script:Cd2Prefs['SiteCode'] | Should -Be 'CD2'
        $script:Cd2Prefs['SMSProvider'] | Should -Be ''
        $script:Cd2Prefs['SMSProvider'] | Should -Not -Match 'example.com'
        $tool = Join-Path $script:DataRoot 'sites\CD2\tools\deployment-helper'
        $script:Cd2.AuditLog | Should -Be (Join-Path $tool 'deployment-audit.jsonl')
        Test-Path -LiteralPath (Join-Path $tool 'rings\AB1-plan.json') | Should -BeFalse
        @(Get-ChildItem -LiteralPath $tool -Recurse -File | Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -match 'AB1' -and $_.Name -ne 'migrated.json' -and $_.Name -ne 'prefs.json' }) | Should -BeNullOrEmpty
    }

    It 'writes the CD2 data only under sites\CD2 and leaves AB1 and the old folder alone' {
        $cd2Root = Join-Path $script:DataRoot 'sites\CD2'
        (Get-Content -LiteralPath (Join-Path $cd2Root 'tools\deployment-helper\deployment-audit.jsonl') -Raw) | Should -Match 'CD2 app'
        (Get-Content -LiteralPath (Join-Path $cd2Root 'tools\deployment-helper\deployment-audit.jsonl') -Raw) | Should -Not -Match 'AB1'
        Test-Path -LiteralPath (Join-Path $cd2Root 'tools\deployment-helper\rings\runs\run-cd2.json') | Should -BeTrue
        $ab1Now = Get-TreeHash -Root (Join-Path $script:DataRoot 'sites\AB1')
        foreach ($k in $script:Ab1After.Keys) {
            if ($k -like '*prefs.json') { continue }
            $ab1Now[$k] | Should -Be $script:Ab1After[$k] -Because "$k of AB1 must not change while CD2 runs"
        }
        foreach ($k in $script:OldHash.Keys) {
            if ($k -like '*DeploymentHelper.prefs.json') { continue }
            (Get-FileHash -LiteralPath $k -Algorithm SHA256).Hash | Should -Be $script:OldHash[$k]
        }
        Test-Path -LiteralPath (Join-Path $script:ToolRoot 'Logs\deployment-audit.jsonl') | Should -BeTrue
        (Get-Content -LiteralPath (Join-Path $script:ToolRoot 'Logs\deployment-audit.jsonl') -Raw) | Should -Not -Match 'CD2'
    }

    It 'gives the AB1 values back when the tool starts for AB1 again' {
        $script:Ab1Again.Site.SiteFolder | Should -Be 'AB1'
        $script:Ab1Again.Prefs['SiteCode'] | Should -Be 'AB1'
        $script:Ab1Again.Prefs['SMSProvider'] | Should -Be 'cm02.example.com' -Because 'AB1 saved this provider before CD2 started'
        (Get-Content -LiteralPath $script:Ab1Again.AuditLog -Raw) | Should -Match 'New AB1 app'
        (Get-Content -LiteralPath $script:Ab1Again.AuditLog -Raw) | Should -Not -Match 'CD2'
        Test-Path -LiteralPath (Join-Path $script:Ab1Again.RingRuns 'run-ab1-new.json') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:Ab1Again.RingRuns 'run-cd2.json') | Should -BeFalse
    }

    It 'keeps a different provider for each site' {
        (Get-Content -LiteralPath (Join-Path $script:DataRoot 'sites\CD2\tools\deployment-helper\prefs.json') -Raw | ConvertFrom-Json).SMSProvider | Should -Be 'cm03.example.com'
        (Get-Content -LiteralPath (Join-Path $script:DataRoot 'sites\AB1\tools\deployment-helper\prefs.json') -Raw | ConvertFrom-Json).SMSProvider | Should -Be 'cm02.example.com'
    }
}

Describe 'Deployment Helper with no site chosen' {
    BeforeAll {
        $script:Saved = @{}
        foreach ($n in $script:EnvNames) { $script:Saved[$n] = [Environment]::GetEnvironmentVariable($n); [Environment]::SetEnvironmentVariable($n, $null) }
        $temp = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $script:Suite = Join-Path $temp 'AppPackagerSuite'
        $script:Data = Join-Path $temp 'Data'
        [void](New-Item -ItemType Directory -Path $script:Suite, $script:Data -Force)
        [Environment]::SetEnvironmentVariable('APP_PACKAGER_DATA_ROOT', $script:Data)
        $script:ToolRoot = New-OldToolFolder -Suite $script:Suite -NoSiteCode
        $script:OldHash = Get-TreeHash -Root $script:ToolRoot
        $script:None = Start-Tool -Folder ''
    }
    AfterAll {
        foreach ($n in $script:EnvNames) { [Environment]::SetEnvironmentVariable($n, $script:Saved[$n]) }
    }

    It 'is not active and uses the old paths beside the script' {
        $script:None.Site.Active | Should -BeFalse
        $script:None.AuditLog | Should -Be (Join-Path $script:ToolRoot 'Logs\deployment-audit.jsonl')
        $script:None.RingPlans | Should -Be (Join-Path $script:ToolRoot 'Rings')
        $script:None.RingRuns | Should -Be (Join-Path $script:ToolRoot 'Rings\runs')
        $script:None.Rollouts | Should -Be (Join-Path $script:ToolRoot 'Rollouts')
    }

    It 'creates no site folder and changes no old file' {
        Test-Path -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites') | Should -BeFalse
        $now = Get-TreeHash -Root $script:ToolRoot
        foreach ($k in $script:OldHash.Keys) { $now[$k] | Should -Be $script:OldHash[$k] }
    }

    It 'saves the site keys and the other keys to the old prefs file' {
        $global:Prefs['SiteCode'] = 'EF3'
        $global:Prefs['SMSProvider'] = 'cm04.example.com'
        Save-DhPreferences -Prefs $global:Prefs
        $p = Get-Content -LiteralPath (Join-Path $script:ToolRoot 'DeploymentHelper.prefs.json') -Raw | ConvertFrom-Json
        $p.SiteCode | Should -Be 'EF3'
        $p.SMSProvider | Should -Be 'cm04.example.com'
        Test-Path -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites') | Should -BeFalse
    }
}

Describe 'Store paths of the user' {
    BeforeAll {
        $script:Site = [pscustomobject]@{ Active = $true; ToolFolder = (Join-Path $TestDrive 'sites\AB1\tools\deployment-helper') }
        $script:Root = Join-Path $TestDrive 'old-tool'
    }

    It 'uses a path that the user chose, with or without a site' {
        Resolve-DhStorePath -Store RingRuns -ToolRoot $script:Root -Pref '\\files.example.com\share\runs' -Site $script:Site | Should -Be '\\files.example.com\share\runs'
        Resolve-DhStorePath -Store AuditLog -ToolRoot $script:Root -Pref 'D:\audit\a.jsonl' -Site $null | Should -Be 'D:\audit\a.jsonl'
    }

    It 'treats the old default under the tool folder as empty, in any letter case and with a trailing slash' {
        Resolve-DhStorePath -Store RingRuns -ToolRoot $script:Root -Pref ((Join-Path $script:Root 'RINGS\RUNS') + '\') -Site $script:Site | Should -Be (Join-Path $script:Site.ToolFolder 'rings\runs')
        Resolve-DhStorePath -Store Rollouts -ToolRoot $script:Root -Pref (Join-Path $script:Root 'Rollouts') -Site $script:Site | Should -Be (Join-Path $script:Site.ToolFolder 'rollouts')
        Resolve-DhStorePath -Store AuditLog -ToolRoot $script:Root -Pref (Join-Path $script:Root 'Logs\deployment-audit.jsonl') -Site $null | Should -Be (Join-Path $script:Root 'Logs\deployment-audit.jsonl')
    }

    It 'returns an empty value for a blank or stale preference' {
        ConvertTo-DhPathPref -Value '  ' -LegacyPath 'C:\x' | Should -Be ''
        ConvertTo-DhPathPref -Value $null -LegacyPath 'C:\x' | Should -Be ''
        ConvertTo-DhPathPref -Value 'C:\x\' -LegacyPath 'C:\x' | Should -Be ''
        ConvertTo-DhPathPref -Value 'C:\y' -LegacyPath 'C:\x' | Should -Be 'C:\y'
    }

    It 'refuses nothing and creates the site folder only with -Create' {
        $p = Resolve-DhStorePath -Store RingPlans -ToolRoot $script:Root -Site $script:Site
        Test-Path -LiteralPath $p | Should -BeFalse
        $q = Resolve-DhStorePath -Store RingPlans -ToolRoot $script:Root -Site $script:Site -Create
        Test-Path -LiteralPath $q | Should -BeTrue
    }
}

Describe 'Static guard: no site store is built from the script folder' {
    It 'builds no audit log, ring or rollout path from $PSScriptRoot in the start script' {
        $hits = [regex]::Matches($script:ShellText, "PSScriptRoot\s+'(Logs\\deployment-audit|Rings|Rollouts)")
        @($hits | ForEach-Object { $_.Value }) | Should -BeNullOrEmpty
    }

    It 'stores no computed default path in the prefs defaults' {
        $script:ShellText | Should -Not -Match "DeploymentAuditLogPath\s*=\s*Join-Path"
        $script:ShellText | Should -Not -Match "RingRunStatePath\s*=\s*Join-Path"
        $script:ShellText | Should -Not -Match "RolloutRunPath\s*=\s*Join-Path"
    }

    It 'sends every audit log, ring and rollout path through Resolve-DhStorePath' {
        foreach ($s in 'AuditLog', 'RingPlans', 'RingRuns', 'Rollouts') {
            $script:ShellText | Should -Match ("Resolve-DhStorePath -Store $s ") -Because $s
        }
    }
}
