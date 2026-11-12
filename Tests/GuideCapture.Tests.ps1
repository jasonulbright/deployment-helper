#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Structure tests of the capture tool. They parse the scripts and run nothing on a lab computer.

BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:ToolPath = Join-Path $script:RepoRoot 'tools\guide-capture\Invoke-GuideCapture.ps1'
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ToolPath, [ref]$null, [ref]$null)
    $script:Text = Get-Content -LiteralPath $script:ToolPath -Raw
}

Describe 'Invoke-GuideCapture' {
    It 'parses without an error' {
        $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($script:ToolPath, [ref]$null, [ref]$errors)
        @($errors) | Should -BeNullOrEmpty
    }

    It 'takes the lock file with New-Item and -ErrorAction Stop, and answers busy on an IOException' {
        $commands = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'New-Item' -and $n.Extent.Text -match 'ItemType\s+File' }, $true))
        $commands.Count | Should -BeGreaterOrEqual 1
        $commands[0].Extent.Text | Should -Match '-ErrorAction\s+Stop'
        $catches = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CatchClauseAst] -and @($n.CatchTypes | Where-Object { $_.TypeName.FullName -eq 'System.IO.IOException' }).Count -gt 0 }, $true))
        $catches.Count | Should -BeGreaterOrEqual 1
        $catches[0].Extent.Text | Should -Match "return 'busy'"
    }

    It 'removes a lock file that is older than 60 minutes' {
        $script:Text | Should -Match 'TotalMinutes\s+-gt\s+60'
    }

    It 'never deletes C:\guide itself, only its own folder and the lock file' {
        $removals = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Remove-Item' }, $true))
        $removals.Count | Should -BeGreaterThan 0
        foreach ($r in $removals) { $r.Extent.Text | Should -Not -Match "C:\\guide['""\s]" }
        $script:Text | Should -Match '\$root = "C:\\guide\\\$app"'
        $script:Text | Should -Match "'C:\\guide\\capture\.lock'"
    }

    It 'uses its own task name and stops the task when the wait ends' {
        $script:Text | Should -Match 'GuideShots-\$app'
        ([regex]::Matches($script:Text, 'Stop-ScheduledTask')).Count | Should -BeGreaterOrEqual 2
    }

    It 'names the app in the lab by a constant, so a worktree uses the same folder' {
        $script:Text | Should -Match '\$app = ''deployment-helper'''
        $script:Text | Should -Not -Match 'Split-Path\s+\$repo\s+-Leaf'
    }

    It 'excludes .claude from the files that it copies' {
        $script:Text | Should -Match "'\.claude'"
    }
}
