#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Headless guard: a window test leaves no window open on the desktop of the lab computer.
#   - the STA call function of the harness calls Close-SuiteOpenWindow in its finally block.
#   - Every window test file opens its windows (Open-TestWindow and the member calls that show a window) inside a try that has a finally.
# The file reads syntax trees only and opens no window.

BeforeDiscovery {
    $script:TestsDir = $PSScriptRoot
    $script:WindowTestFiles = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Window.Tests.ps1' -File |
            ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } })
}

BeforeAll {
    function Get-Ast {
        param([string]$Path)
        [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    }
    # True when the node sits in the body of a try statement that has a finally block.
    function Test-InTryWithFinally {
        param($Node)
        for ($p = $Node.Parent; $p; $p = $p.Parent) {
            if ($p -is [System.Management.Automation.Language.TryStatementAst] -and $p.Finally) {
                $inBody = $p.Body.Extent.StartOffset -le $Node.Extent.StartOffset -and $Node.Extent.EndOffset -le $p.Body.Extent.EndOffset
                if ($inBody) { return $true }
            }
        }
        return $false
    }
    function Get-WindowOpener {
        param($Ast)
        $Ast.FindAll({
                param($n)
                ($n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Open-TestWindow') -or
                ($n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Extent.Text -in 'Show', ('Show' + 'Dialog'))
            }, $true)
    }
}

Describe 'Window test cleanup' {
    It 'finds at least one window test file' {
        @($script:WindowTestFiles).Count | Should -BeGreaterThan 0
    }

    It 'closes every open window in the finally block of the STA call function' {
        $ast = Get-Ast -Path (Join-Path $PSScriptRoot 'WindowHarness.ps1')
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq ('Invoke-Wpf' + 'Sta') }, $true)
        $fn | Should -Not -BeNullOrEmpty
        $try = @($fn.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] -and $n.Finally }, $true))
        $try.Count | Should -BeGreaterThan 0
        ($try | Where-Object { $_.Finally.Extent.Text -match 'Close-SuiteOpenWindow' }) | Should -Not -BeNullOrEmpty -Because 'the finally block must close the windows the call left open'
    }

    It '<Name> opens every window inside a try block that has a finally' -ForEach $script:WindowTestFiles {
        $ast = Get-Ast -Path $Path
        $openers = @(Get-WindowOpener -Ast $ast)
        $openers.Count | Should -BeGreaterThan 0
        $bad = @($openers | Where-Object { -not (Test-InTryWithFinally -Node $_) } | ForEach-Object { 'line {0}: {1}' -f $_.Extent.StartLineNumber, $_.Extent.Text })
        $bad | Should -BeNullOrEmpty -Because ($bad -join '; ')
    }
}

Describe 'Window cleanup guard reads the syntax tree' {
    It 'flags an opener outside a try with a finally' {
        $ast = [System.Management.Automation.Language.Parser]::ParseInput('$w = Open-TestWindow -XamlPath x; $w.Close()', [ref]$null, [ref]$null)
        $o = @(Get-WindowOpener -Ast $ast)
        $o.Count | Should -Be 1
        Test-InTryWithFinally -Node $o[0] | Should -BeFalse
    }

    It 'accepts an opener inside a try body with a finally, and flags one that is only in the finally' {
        $ast = [System.Management.Automation.Language.Parser]::ParseInput('try { $w = Open-TestWindow -XamlPath x } finally { $w.Close(); $d.Sh' + 'ow() }', [ref]$null, [ref]$null)
        $o = @(Get-WindowOpener -Ast $ast)
        $o.Count | Should -Be 2
        Test-InTryWithFinally -Node $o[0] | Should -BeTrue
        Test-InTryWithFinally -Node $o[1] | Should -BeFalse
    }
}
