#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }

<#
.SYNOPSIS
    Pester 5.x window test for the compact 1080p density and for vacant space. The main window is opened in both themes at
    the work area of the session (1920x1032 on a 1080p screen at 100 percent) and at 1024x720, and measured: sidebar
    buttons, inline buttons, single-line boxes, grid rows and headers, the width that the content uses, and the vacant area of
    each view (the part of the view below or right of its last content element). The 1024x720 case proves that nothing clips.
    The visible run prints the work area of the session (SystemParameters.WorkArea).
#>

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed -and -not $env:SUITE_WINDOW_GUARD_REPORTED) { Write-Host ('Window tests skipped: {0}.' -f $script:WindowGuard.Reason) }
    $script:Cases = @(foreach ($dark in $true, $false) {
            $design = Get-DesignSize
            foreach ($size in @(@(1024, 720), @($design.W, $design.H))) {
                @{ Dark = $dark; Theme = $(if ($dark) { 'dark' } else { 'light' }); W = $size[0]; H = $size[1]; Size = ('{0}x{1}' -f $size[0], $size[1]) }
            }
        })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:Design = Get-DesignSize
    Add-Type -AssemblyName PresentationFramework
    $wa = [System.Windows.SystemParameters]::WorkArea
    Write-Host ('Work area of this session: {0}x{1}' -f $wa.Width, $wa.Height)
    $script:Measured = @{}
    foreach ($case in @(
            @{ Dark = $true; W = 1024; H = 720 }, @{ Dark = $true; W = $script:Design.W; H = $script:Design.H },
            @{ Dark = $false; W = 1024; H = 720 }, @{ Dark = $false; W = $script:Design.W; H = $script:Design.H })) {
        $key = '{0}|{1}x{2}' -f $(if ($case.Dark) { 'dark' } else { 'light' }), $case.W, $case.H
        $case.Config = $script:AppConfig
        $script:Measured[$key] = Invoke-WpfSta -Arguments $case -Script {
            $cfg = $Arguments.Config
            $w = $null
            try {
                $w = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark -Width $Arguments.W -Height $Arguments.H
                Initialize-ShellTheme -Window $w -Dark $Arguments.Dark -WorkflowButtonNames $cfg.WorkflowButtons -OptionsButtonNames $cfg.OptionsButtons -LogLabelName $cfg.LogLabel
                Wait-Layout $w
                $sidebar = @(foreach ($n in @($cfg.WorkflowButtons) + @($cfg.OptionsButtons)) { $b = $w.FindName($n); $p = $b.TransformToAncestor($w).Transform([System.Windows.Point]::new(0, 0)); [pscustomobject]@{ Name = $n; Top = $p.Y; Bottom = $p.Y + $b.ActualHeight; Height = $b.ActualHeight } })
                $views = @(foreach ($view in @(@{ Name = 'Deployment'; Show = $cfg.FormPane; Hide = $cfg.RingPane }, @{ Name = 'Ring Deployment'; Show = $cfg.RingPane; Hide = $cfg.FormPane })) {
                        $w.FindName($view.Show).Visibility = 'Visible'; $w.FindName($view.Hide).Visibility = 'Collapsed'
                        Wait-Layout $w
                        $pane = $w.FindName($view.Show)
                        $buttons = @(); $boxes = @(); $grids = @(); $bottom = 0.0; $right = 0.0; $spans = @()
                        foreach ($e in (Get-VisualDescendants $pane)) {
                            if ($e -isnot [System.Windows.FrameworkElement] -or $e.ActualWidth -le 0 -or $e.ActualHeight -le 0 -or -not (Test-VisibleInTree $e)) { continue }
                            $isContent = ($e -is [System.Windows.Controls.TextBlock] -and -not [string]::IsNullOrWhiteSpace($e.Text)) -or $e -is [System.Windows.Controls.Primitives.ButtonBase] -or $e -is [System.Windows.Controls.TextBox] -or $e -is [System.Windows.Controls.ComboBox] -or $e -is [System.Windows.Controls.DataGrid]
                            $p = $e.TransformToAncestor($pane).Transform([System.Windows.Point]::new(0, 0))
                            if ($isContent) { $bottom = [math]::Max($bottom, $p.Y + $e.ActualHeight); $right = [math]::Max($right, $p.X + $e.ActualWidth); $spans += , @($p.Y, ($p.Y + $e.ActualHeight)) }
                            if ($e -is [System.Windows.Controls.Button] -and $e.Name) { $buttons += [pscustomobject]@{ Name = $e.Name; Height = $e.ActualHeight } }
                            if ($e -is [System.Windows.Controls.TextBox] -and $e.Name -in 'txtTargetName', 'txtCollection', 'txtRingObject') { $boxes += [pscustomobject]@{ Name = $e.Name; Height = $e.ActualHeight } }
                            if ($e -is [System.Windows.Controls.DataGrid]) {
                                $hdr = $null; foreach ($dgx in (Get-VisualDescendants $e)) { if ($dgx -is [System.Windows.Controls.Primitives.DataGridColumnHeadersPresenter]) { $hdr = $dgx; break } }
                                $row = $null; foreach ($dgx in (Get-VisualDescendants $e)) { if ($dgx -is [System.Windows.Controls.DataGridRow]) { $row = $dgx; break } }
                                $grids += [pscustomobject]@{ Name = $e.Name; Width = $e.ActualWidth; HeaderHeight = $(if ($hdr) { $hdr.ActualHeight } else { $null }); RowHeight = $(if ($row) { $row.ActualHeight } else { $null }) }
                            }
                        }
                        # The largest empty band between two content elements (the free height inside the view, not only below it).
                        $gap = 0.0; $cover = 0.0
                        foreach ($sp in @($spans | Sort-Object { $_[0] })) { if ($sp[0] -gt $cover) { $gap = [math]::Max($gap, $sp[0] - $cover) }; $cover = [math]::Max($cover, $sp[1]) }
                        [pscustomobject]@{ View = $view.Name; Gap = $gap; Width = $pane.ActualWidth; Height = $pane.ActualHeight; VacantHeight = $pane.ActualHeight - $bottom; VacantWidth = $pane.ActualWidth - $right; Buttons = $buttons; Boxes = $boxes; Grids = $grids }
                    })
                [pscustomobject]@{ Width = $w.ActualWidth; Height = $w.ActualHeight; Sidebar = $sidebar; Views = $views }
            }
            finally { if ($w) { $w.Close() } }
        }
    }
}

AfterAll {
    if ($script:WindowGuard.Allowed) { Close-WpfSession }
}

Describe 'Density <Theme> <Size>' -ForEach $script:Cases -Skip:(-not $script:WindowGuard.Allowed) {
    It 'sizes every sidebar button between 34 and 36 px and fits all of them without scrolling' {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size]
        foreach ($b in $m.Sidebar) { $b.Height | Should -BeIn 34, 35, 36 -Because $b.Name }
        ($m.Sidebar | Select-Object -Last 1).Bottom | Should -BeLessThan $m.Height
    }

    It 'sizes the inline buttons and the single-line boxes at 26 px in the <View> view' -ForEach @(@{ View = 'Deployment' }, @{ View = 'Ring Deployment' }) {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size].Views | Where-Object View -eq $View
        foreach ($b in $m.Buttons | Where-Object { $_.Name -match '^btn(Browse|Ring(Browse|Reload|Open|Expand|Refresh))' }) { $b.Height | Should -BeIn 26, 27, 28 -Because $b.Name }
        foreach ($b in $m.Boxes) { $b.Height | Should -BeIn 26, 27, 28 -Because $b.Name }
    }

    It 'gives grid rows 24 to 26 px and headers 26 px in the <View> view' -ForEach @(@{ View = 'Ring Deployment' }) {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size].Views | Where-Object View -eq $View
        foreach ($g in $m.Grids) {
            if ($null -ne $g.HeaderHeight) { $g.HeaderHeight | Should -BeLessOrEqual 26.5 -Because $g.Name }
            if ($null -ne $g.RowHeight) { [math]::Round($g.RowHeight) | Should -BeIn 24, 25, 26 -Because $g.Name }
        }
    }

    It 'lets the content fill the width in the <View> view' -ForEach @(@{ View = 'Deployment' }, @{ View = 'Ring Deployment' }) {
        $mm = $script:Measured['{0}|{1}' -f $Theme, $Size]
        $m = $mm.Views | Where-Object View -eq $View
        $m.Width | Should -BeGreaterOrEqual ($mm.Width - 300)
        foreach ($g in $m.Grids) { $g.Width | Should -BeGreaterOrEqual ($m.Width * 0.3) -Because $g.Name }
    }
}

Describe 'Vacant space <Theme> <Size>' -ForEach $script:Cases -Skip:(-not $script:WindowGuard.Allowed) {
    It 'leaves no vacant area of more than 160 px wide or 120 px high in any view that has content to fill it' {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size]
        $small = ($Size -eq '1024x720')
        $bad = @()
        foreach ($v in $m.Views) {
            if ($v.VacantHeight -gt 120 -and -not $small) { $bad += ('{0}: {1} px vacant below the last element (view {2}x{3})' -f $v.View, [math]::Round($v.VacantHeight), [math]::Round($v.Width), [math]::Round($v.Height)) }
            if ($v.Gap -gt 120 -and -not $small) { $bad += ('{0}: {1} px empty band between two content elements' -f $v.View, [math]::Round($v.Gap)) }
            if ($v.VacantWidth -gt 160 -and -not $small) { $bad += ('{0}: {1} px vacant right of the last element (view {2}x{3})' -f $v.View, [math]::Round($v.VacantWidth), [math]::Round($v.Width), [math]::Round($v.Height)) }
            # The view scrolls vertically in a small window, so only the width must fit.
            if ($small -and $v.VacantWidth -lt -2) { $bad += ('{0}: content clips by {1} px on the right' -f $v.View, [math]::Round(-$v.VacantWidth)) }
        }
        $bad | Should -BeNullOrEmpty -Because ($bad -join '; ')
    }
}
