#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Window test for the hover, focus and pressed states of the buttons of the main window, in both themes.
# It needs the real mouse, so it runs only in a visible run (SUITE_VISIBLE_WINDOWS=1, the labadmin session of CLIENT01).
# The measurement is Get-SuiteButtonStateReport of SuiteCommon. For each sampled button and each state it requires:
#   - the button text has 4.5:1 contrast on the button fill (the backdrop of the button where the fill is transparent)
#   - the button is told apart from its backdrop: its border or its fill has 3:1 contrast on it

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $guard = Test-WindowTestsAllowed
    $script:Visible = ($env:SUITE_VISIBLE_WINDOWS -eq '1')
    if ($guard.Allowed -and -not $script:Visible) { Write-Host 'Button state tests skipped: SUITE_VISIBLE_WINDOWS is not 1 (they need the real mouse).' }
    # The window guard of this file also needs the visible run, because the test moves the real mouse.
    $script:WindowGuard = [pscustomobject]@{ Allowed = ($guard.Allowed -and $script:Visible); Reason = $guard.Reason }
    $script:ThemeCases = @(@{ Dark = $true; Theme = 'dark' }, @{ Dark = $false; Theme = 'light' })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    $guard = Test-WindowTestsAllowed
    $script:Visible = ($env:SUITE_VISIBLE_WINDOWS -eq '1')
    $script:WindowGuard = [pscustomobject]@{ Allowed = ($guard.Allowed -and $script:Visible); Reason = $guard.Reason }
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    # One dialog with the primary (blue) and the secondary buttons: the DP group picker, built from its XAML text the way the script builds it.
    $shellPath = Join-Path (Split-Path -Parent $PSScriptRoot) $script:AppConfig.Shell
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($shellPath, [ref]$null, [ref]$null)
    $dialogXaml = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow' -and $n.Value -match 'Title="Select DP Groups"' }, $true) | ForEach-Object { $_.Value })
    $oneDialog = if ($dialogXaml) { $dialogXaml[0].Replace($script:SharedStyleMarker, (Get-SharedStyleXaml)) } else { $null }
    $script:Measured = @{}
    foreach ($dark in $true, $false) {
        $key = if ($dark) { 'dark' } else { 'light' }
        $script:Measured[$key] = Invoke-WpfSta -Arguments @{ Dark = $dark; Config = $script:AppConfig; Dialog = $oneDialog } -Script {
            $cfg = $Arguments.Config
            $w = $null
            try {
            $w = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark
            Initialize-ShellTheme -Window $w -Dark $Arguments.Dark -WorkflowButtonNames $cfg.WorkflowButtons -OptionsButtonNames $cfg.OptionsButtons -LogLabelName $cfg.LogLabel
            $w.FindName($cfg.FormPane).Visibility = 'Visible'; $w.FindName($cfg.RingPane).Visibility = 'Collapsed'
            Wait-Layout $w
            Get-SuiteButtonStateReport -Window $w -Dark $Arguments.Dark | ForEach-Object { $_ | Add-Member -NotePropertyName Source -NotePropertyValue 'main' -PassThru }
            if ($Arguments.Dialog) {
                [xml]$x = $Arguments.Dialog
                $d = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
                try {
                    $d.WindowStartupLocation = 'Manual'; $d.Left = 120; $d.Top = 120; $d.ShowInTaskbar = $false
                    Set-DialogTheme -Dialog $d
                    $d.Show(); Wait-Layout $d
                    Get-SuiteButtonStateReport -Window $d -Dark $Arguments.Dark -Max 3 | ForEach-Object { $_ | Add-Member -NotePropertyName Source -NotePropertyValue 'dialog' -PassThru }
                }
                finally { $d.Close() }
            }
            }
            finally { if ($w) { $w.Close() } }
        }
    }
}

Describe 'Button states <Theme>' -ForEach $script:ThemeCases -Skip:(-not $script:WindowGuard.Allowed) {
    BeforeAll { $script:Rows = @($script:Measured[$Theme] | ForEach-Object { $_ } | Where-Object { $_.PSObject.Properties['States'] }) }

    It 'samples at least 3 buttons and reaches the focus, hover and pressed states' {
        $script:Rows.Count | Should -BeGreaterOrEqual 3
        $bad = @($script:Rows | Where-Object { -not $_.Hover -or -not $_.Pressed -or -not $_.Focused } | ForEach-Object { '{0}: hover {1} pressed {2} focused {3}' -f $_.Name, $_.Hover, $_.Pressed, $_.Focused })
        $bad | Should -BeNullOrEmpty -Because ('the mouse and the keyboard focus must reach the button, or the states are not measured: ' + ($bad -join '; '))
    }

    It 'keeps 4.5:1 text contrast in every state' {
        $bad = @()
        foreach ($r in $script:Rows) { foreach ($s in $r.States.Keys) { if ($r.States[$s].Text -lt 4.5) { $bad += ('{0} {1}: {2}' -f $r.Name, $s, $r.States[$s].Text) } } }
        $bad | Should -BeNullOrEmpty -Because ($bad -join '; ')
    }

    It 'tells the button from its backdrop (3:1) in every state' {
        $bad = @()
        foreach ($r in $script:Rows) { foreach ($s in $r.States.Keys) { if ($r.States[$s].Edge -lt 3.0) { $bad += ('{0} {1}: {2}' -f $r.Name, $s, $r.States[$s].Edge) } } }
        $bad | Should -BeNullOrEmpty -Because ($bad -join '; ')
    }

    It 'keeps a blue button blue in hover and pressed and never black; a dialog button moves one visible step in each state' {
        $bad = @()
        foreach ($r in $script:Rows) {
            if (-not $r.States['normal'].Blue) { continue }
            foreach ($s in 'hover', 'pressed') {
                $st = $r.States[$s]
                if (-not $st.Blue) { $bad += ('{0} {1} {2}: fill {3} is not blue' -f $r.Source, $r.Name, $s, $st.Fill) }
                if ($st.FillLuminance -lt 0.03) { $bad += ('{0} {1} {2}: fill {3} luminance {4} is black' -f $r.Source, $r.Name, $s, $st.Fill, $st.FillLuminance) }
                $step = ([math]::Max($st.FillLuminance, $r.States['normal'].FillLuminance) + 0.05) / ([math]::Min($st.FillLuminance, $r.States['normal'].FillLuminance) + 0.05)
                if ($r.Source -eq 'dialog' -and $step -lt 1.05) { $bad += ('{0} {1} {2}: fill {3} is the same as normal {4}' -f $r.Source, $r.Name, $s, $st.Fill, $r.States['normal'].Fill) }
            }
        }
        $bad | Should -BeNullOrEmpty -Because ($bad -join '; ')
    }

    It 'never flips a neutral button to black or to white in hover and pressed' {
        $bad = @()
        foreach ($r in $script:Rows) {
            if ($r.States['normal'].Blue) { continue }
            $n = $r.States['normal'].FillLuminance
            foreach ($s in 'hover', 'pressed') {
                $l = $r.States[$s].FillLuminance
                $blue = $r.States[$s].Blue
                if (-not $blue -and (($l -lt 0.03 -and $n -gt 0.1) -or ($l -gt 0.9 -and $n -lt 0.1))) { $bad += ('{0} {1} {2}: fill {3} flips from {4}' -f $r.Source, $r.Name, $s, $r.States[$s].Fill, $r.States['normal'].Fill) }
            }
        }
        $bad | Should -BeNullOrEmpty -Because ($bad -join '; ')
    }

    It 'measures a button of the main window and the accent and neutral buttons of a dialog' {
        @($script:Rows | Where-Object { $_.Source -eq 'main' }).Count | Should -BeGreaterOrEqual 3
        @($script:Rows | Where-Object { $_.Source -eq 'dialog' }).Count | Should -BeGreaterOrEqual 2
    }

    It 'lists the measured values' {
        foreach ($r in $script:Rows) { Write-Host ('{0} {1}: ' -f $Theme, $r.Name) -NoNewline; Write-Host (($r.States.Keys | ForEach-Object { '{0} text {1} edge {2}' -f $_, $r.States[$_].Text, $r.States[$_].Edge }) -join ' | ') }
        $script:Rows.Count | Should -BeGreaterThan 0
    }
}
