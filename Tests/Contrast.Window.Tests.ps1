#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Window test for the disabled state of the controls of the main window, in both themes.
# A disabled (inactive) control is exempt from the 4.5:1 text rule: WCAG 1.4.3 exempts inactive components.
# So no contrast target applies to disabled text or a disabled border. A disabled control must still look dimmed,
# so that it reads as disabled, and its label must stay readable:
#   - EnabledText / DisabledText >= 1.25 (a disabled button that looks like an enabled one is a defect)
#   - DisabledText >= 1.5 (below that the label vanishes)
# The measurement is Get-InteractiveContrast of WindowMeasure.ps1: every visible named button, text box and combo box,
# measured enabled and then disabled.

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed -and -not $env:SUITE_WINDOW_GUARD_REPORTED) { Write-Host ('Window tests skipped: {0}.' -f $script:WindowGuard.Reason) }
    $script:ThemeCases = @(@{ Dark = $true; Theme = 'dark' }, @{ Dark = $false; Theme = 'light' })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:Measured = @{}
    foreach ($dark in $true, $false) {
        $key = if ($dark) { 'dark' } else { 'light' }
        $script:Measured[$key] = Invoke-WpfSta -Arguments @{ Dark = $dark; Config = $script:AppConfig } -Script {
            $cfg = $Arguments.Config
            $w = $null
            try {
                $w = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark
                Initialize-ShellTheme -Window $w -Dark $Arguments.Dark -WorkflowButtonNames $cfg.WorkflowButtons -OptionsButtonNames $cfg.OptionsButtons -LogLabelName $cfg.LogLabel
                $rows = @()
                foreach ($pane in @(@($cfg.FormPane, $cfg.RingPane), @($cfg.RingPane, $cfg.FormPane), @($cfg.RolloutPane, $cfg.RingPane))) {
                    $w.FindName($pane[0]).Visibility = 'Visible'; $w.FindName($pane[1]).Visibility = 'Collapsed'
                    Wait-Layout $w
                    $rows += @(Get-InteractiveContrast -Window $w)
                }
                $rows
            }
            finally { if ($w) { $w.Close() } }
        }
    }
}

AfterAll {
    if ($script:WindowGuard.Allowed) { Close-WpfSession }
}

Describe 'Disabled controls <Theme>' -ForEach $script:ThemeCases -Skip:(-not $script:WindowGuard.Allowed) {
    BeforeAll { $script:Rows = @($script:Measured[$Theme] | ForEach-Object { $_ } | Where-Object { $_.PSObject.Properties['DisabledText'] }) }

    It 'measures the buttons, text boxes and combo boxes of both panes' {
        $script:Rows.Count | Should -BeGreaterThan 5
    }

    It 'shows every disabled control clearly dimmed (enabled text over disabled text at least 1.25) and still readable (at least 1.5)' {
        # One documented limit: in the light theme SuiteCommon sets a local white foreground on the blue sidebar buttons, so a
        # disabled sidebar button would not dim. The app never disables a sidebar button (Theme.Guard.Tests.ps1 asserts both
        # facts). The fix belongs in SuiteCommon. See KnownThemeLimits.ps1, sidebar-disabled-white-on-blue.
        . "$PSScriptRoot\KnownThemeLimits.ps1"
        . "$PSScriptRoot\WindowConfig.ps1"
        $limit = @($script:KnownThemeLimit | Where-Object { $_.Id -eq 'sidebar-disabled-white-on-blue' })
        $limit.Count | Should -Be 1
        $sidebar = @($script:AppConfig.WorkflowButtons) + @($script:AppConfig.OptionsButtons)
        $bad = @()
        foreach ($r in $script:Rows) {
            if ($null -eq $r.EnabledText -or $null -eq $r.DisabledText) { continue }
            if ($Theme -eq 'light' -and $sidebar -contains $r.Name) { continue }
            $dim = $r.EnabledText / [math]::Max($r.DisabledText, 0.01)
            if ($dim -lt 1.25) { $bad += ('{0}: enabled {1}, disabled {2}, dimming {3}' -f $r.Name, [math]::Round($r.EnabledText, 2), [math]::Round($r.DisabledText, 2), [math]::Round($dim, 2)) }
            if ($r.DisabledText -lt 1.5) { $bad += ('{0}: disabled text {1} is not readable' -f $r.Name, [math]::Round($r.DisabledText, 2)) }
        }
        $bad | Should -BeNullOrEmpty -Because ($bad -join '; ')
    }

    It 'lists the measured values' {
        foreach ($r in $script:Rows) { Write-Host ('{0} {1} {2}: enabled text {3} border {4} | disabled text {5} border {6}' -f $Theme, $r.Type, $r.Name, $r.EnabledText, $r.EnabledBorder, $r.DisabledText, $r.DisabledBorder) }
        $script:Rows.Count | Should -BeGreaterThan 0
    }
}
