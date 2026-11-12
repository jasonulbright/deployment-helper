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
        $script:Text | Should -Match '\$stopRunner = .*Stop-ScheduledTask'
        ([regex]::Matches($script:Text, '\$stopRunner')).Count | Should -BeGreaterOrEqual 3
    }

    It 'ends a capture in the order: stop the task and kill the processes, unregister, remove the folder, release the lock' {
        $finally = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] -and $null -ne $n.Finally }, $true)
        $finally | Should -Not -BeNullOrEmpty
        $f = $finally.Finally.Extent.Text
        $kill = $f.IndexOf('$stopRunner')
        $unregister = $f.IndexOf('Unregister-ScheduledTask')
        $folder = $f.IndexOf('Remove-Item $r -Recurse')
        $lock = $f.IndexOf('Remove-Item $l -Force')
        $kill | Should -BeGreaterOrEqual 0
        $unregister | Should -BeGreaterThan $kill
        $folder | Should -BeGreaterThan $unregister
        $lock | Should -BeGreaterThan $folder
    }

    It 'kills the runner and the app by the lab folder in the command line, and uses the step on a timeout' {
        $script:Text | Should -Match 'Get-CimInstance Win32_Process'
        $script:Text | Should -Match 'Stop-Process'
        $script:Text | Should -Match 'CommandLine -like "\*\$r\\\*"'
        $timeout = $script:Text.IndexOf('The capture did not finish')
        $timeout | Should -BeGreaterThan 0
        $script:Text.Substring(0, $timeout).LastIndexOf('$stopRunner') | Should -BeGreaterThan $script:Text.IndexOf('if (-not $done) {')
    }

    It 'renames a stale lock before it deletes it, and treats a lost lock as busy' {
        $script:Text | Should -Match 'Move-Item -LiteralPath \$l -Destination \$old'
        $script:Text.IndexOf('Move-Item -LiteralPath $l') | Should -BeLessThan $script:Text.IndexOf('Remove-Item -LiteralPath $old')
        $script:Text | Should -Match 'if \(-not \$item\) \{ return ''busy'' \}'
    }

    It 'names the app in the lab by a constant, so a worktree uses the same folder' {
        $script:Text | Should -Match '\$app = ''deployment-helper'''
        $script:Text | Should -Not -Match 'Split-Path\s+\$repo\s+-Leaf'
    }

    It 'excludes .claude from the files that it copies' {
        $script:Text | Should -Match "'\.claude'"
    }
}
