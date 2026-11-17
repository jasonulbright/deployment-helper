#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Text tests of the main window XAML and of the dialogs that the start script builds from XAML text.
# They read the files only and open no window, so they run on every host, also where the window tests are off.
# The loaded windows are checked by Window.Tests.ps1. The app-specific names live in WindowConfig.ps1.

BeforeDiscovery {
    . "$PSScriptRoot\WindowConfig.ps1"
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path (Split-Path -Parent $PSScriptRoot) $script:AppConfig.Shell), [ref]$null, [ref]$null)
    $script:NameCases = @($ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'FindName' -and
                $n.Expression.Extent.Text -eq '$window' -and $n.Arguments.Count -eq 1
            }, $true) | ForEach-Object { $_.Arguments[0].Value } | Sort-Object -Unique | ForEach-Object { @{ Name = $_ } })
}

BeforeAll {
    . "$PSScriptRoot\WindowConfig.ps1"
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:ShellPath = Join-Path $script:Root $script:AppConfig.Shell
    $script:XamlText = [System.IO.File]::ReadAllText((Join-Path $script:Root $script:AppConfig.Xaml))
    $script:ShellText = [System.IO.File]::ReadAllText($script:ShellPath)
    $script:ShellAst = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
    $script:WindowNames = @($script:ShellAst.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'FindName' -and
                $n.Expression.Extent.Text -eq '$window' -and $n.Arguments.Count -eq 1
            }, $true) | ForEach-Object { $_.Arguments[0].Value } | Sort-Object -Unique)
    # A dialog XAML string takes the shared styles from MainWindow.xaml, as the start script does at run time.
    $script:DialogXaml = @($script:ShellAst.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow'
            }, $true) | ForEach-Object { $_.Value.Replace($script:SharedStyleMarker, (Get-SharedStyleXaml)) })
}

Describe 'Window XAML text' {
    It 'declares each x:Name once' {
        $names = [regex]::Matches($script:XamlText, 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        ($names | Group-Object | Where-Object Count -gt 1 | ForEach-Object Name) | Should -BeNullOrEmpty
    }

    It 'finds the script look-ups of the window' {
        $script:WindowNames.Count | Should -BeGreaterThan 30
    }

    It 'has the control <Name> that the script looks up' -ForEach $script:NameCases {
        $script:XamlText | Should -Match ('x:Name="{0}"' -f [regex]::Escape($Name))
    }

    It 'has each configured sidebar button, the log label and the two panes' {
        foreach ($n in @($script:AppConfig.WorkflowButtons) + @($script:AppConfig.OptionsButtons) + @($script:AppConfig.LogLabel, $script:AppConfig.FormPane, $script:AppConfig.RingPane)) {
            $script:XamlText | Should -Match ('x:Name="{0}"' -f $n)
        }
    }

    It 'has the title that the configuration names' {
        $script:XamlText | Should -Match ('\sTitle="{0}"' -f [regex]::Escape($script:AppConfig.Title))
    }
}

Describe 'Dialog text' {
    It 'has every control that the dialog code looks up' {
        $lookups = @([regex]::Matches($script:ShellText, "\`$(?:dlg|script:BrowseLoadDialog)\.FindName\('([^']+)'\)") | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $lookups.Count | Should -BeGreaterThan 3
        $all = $script:DialogXaml -join "`n"
        foreach ($n in $lookups) { $all | Should -Match ('x:Name="{0}"' -f [regex]::Escape($n)) }
    }

    It 'gives every dialog in XAML a title that differs from the main window title and from the other dialogs' {
        $titles = @($script:DialogXaml | ForEach-Object { if ($_ -match '\sTitle="([^"]*)"') { $Matches[1] } } | Where-Object { $_ })
        $titles.Count | Should -BeGreaterThan 0
        $titles | Should -Not -Contain $script:AppConfig.Title
        ($titles | Select-Object -Unique).Count | Should -Be $titles.Count
    }

    It 'gives every message box a fixed title that differs from the main window title' {
        $calls = @($script:ShellAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Show-ThemedMessage' }, $true))
        $calls.Count | Should -BeGreaterThan 3
        $wrapper = 0
        $bad = foreach ($c in $calls) {
            $e = $c.CommandElements
            $title = $null
            for ($i = 1; $i -lt $e.Count - 1; $i++) {
                if ($e[$i] -is [System.Management.Automation.Language.CommandParameterAst] -and $e[$i].ParameterName -eq 'Title' -and $e[$i + 1] -is [System.Management.Automation.Language.StringConstantExpressionAst]) { $title = $e[$i + 1].Value }
            }
            # A wrapper function passes its own -Title parameter through. Only a call with a literal title is checked here.
            if (-not $title) { $wrapper++; continue }
            if ($title -eq $script:AppConfig.Title) { 'line {0}: the title "{1}" equals the main window title' -f $c.Extent.StartLineNumber, $title }
        }
        @($bad) | Should -BeNullOrEmpty
        $wrapper | Should -BeLessThan 3 -Because 'a literal title is the rule, a variable title is for the wrapper function only'
    }
}
