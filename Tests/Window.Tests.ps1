#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Window tests. The main window and the dialogs that the start script builds from XAML text
# are opened off screen on an STA thread, in both themes, at 1024x720 and at the work area of the session, and measured:
# the deployment pane and the ring pane. The chrome, the buttons, the contrast and the density of the same windows
# are measured by Chrome, Contrast and Density.Window.Tests.ps1. The app-specific names live in WindowConfig.ps1.
# The tests that open a window run only with the opt-in SUITE_ALLOW_WINDOW_TESTS=1 (WindowHarness.ps1).
# The text tests of the XAML are in WindowXaml.Tests.ps1. They open no window and always run.

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    . "$PSScriptRoot\WindowConfig.ps1"
    # No window opens unless this guard allows it (the opt-in must be set). The skipped tests report the reason.
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed -and -not $env:SUITE_WINDOW_GUARD_REPORTED) { Write-Host ('Window tests skipped: {0}.' -f $script:WindowGuard.Reason) }
    $script:Matrix = foreach ($dark in $true, $false) {
        $design = Get-DesignSize
        foreach ($size in @(@(1024, 720), @($design.W, $design.H))) {
            @{ Dark = $dark; Theme = $(if ($dark) { 'dark' } else { 'light' }); W = $size[0]; H = $size[1]; Size = ('{0}x{1}' -f $size[0], $size[1]) }
        }
    }
    $script:ViewCases = @(@{ View = 'Deployment' }, @{ View = 'Ring Deployment' })
    $shellAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path (Split-Path -Parent $PSScriptRoot) $script:AppConfig.Shell), [ref]$null, [ref]$null)
    $dialogCount = @($shellAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow' }, $true)).Count
    $script:DialogCases = @(for ($i = 0; $i -lt $dialogCount; $i++) { @{ Index = $i } })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    . "$PSScriptRoot\WindowConfig.ps1"
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:ShellPath = Join-Path $script:Root $script:AppConfig.Shell
    $script:XamlPath = Join-Path $script:Root $script:AppConfig.Xaml
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

    $script:Design = $null
    $script:Measured = @{}
    $script:WindowGuard = Test-WindowTestsAllowed
    if ($script:WindowGuard.Allowed) {
        $script:Design = Get-DesignSize
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
                    $out = @()
                    foreach ($view in @(@{ Name = 'Deployment'; Show = $cfg.FormPane; Hide = $cfg.RingPane }, @{ Name = 'Ring Deployment'; Show = $cfg.RingPane; Hide = $cfg.FormPane })) {
                        $w.FindName($view.Show).Visibility = 'Visible'; $w.FindName($view.Hide).Visibility = 'Collapsed'
                        Wait-Layout $w
                        $out += [PSCustomObject]@{
                            View      = $view.Name
                            Clipped   = @(Get-ClippedText $w)
                            Offscreen = @(Get-OffscreenControls $w)
                        }
                    }
                    , $out
                }
                finally { if ($w) { $w.Close() } }
            }
        }
    }
}

AfterAll { if ($script:WindowGuard.Allowed) { Close-WpfSession } }

Describe 'Main window XAML' -Skip:(-not $script:WindowGuard.Allowed) {
    It 'loads as a MetroWindow with the app title and a minimum size that fits 1024x720' {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:AppConfig.Xaml } -Script {
            $w = $null
            try {
                $w = Open-TestWindow -XamlPath (Join-Path $Root $Arguments.Xaml)
                [PSCustomObject]@{ Type = $w.GetType().FullName; Title = $w.Title; MinW = $w.MinWidth; MinH = $w.MinHeight }
            }
            finally { if ($w) { $w.Close() } }
        }
        $r.Type | Should -Be 'MahApps.Metro.Controls.MetroWindow'
        $r.Title | Should -Be $script:AppConfig.Title
        $r.MinW | Should -BeLessOrEqual 1024
        $r.MinH | Should -BeLessOrEqual 720
    }

    It 'finds every control that the script looks up in the loaded window' {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:AppConfig.Xaml; Names = $script:WindowNames } -Script {
            $w = $null
            try {
                $w = Open-TestWindow -XamlPath (Join-Path $Root $Arguments.Xaml)
                , @($Arguments.Names | Where-Object { -not $w.FindName($_) })
            }
            finally { if ($w) { $w.Close() } }
        }
        @($r) | Should -BeNullOrEmpty
    }

    It 'shows the deployment pane first and keeps the ring pane hidden' {
        $r = Invoke-WpfSta -Arguments @{ Config = $script:AppConfig } -Script {
            $w = $null
            try {
                $w = Open-TestWindow -XamlPath (Join-Path $Root $Arguments.Config.Xaml)
                [PSCustomObject]@{ Form = [string]$w.FindName($Arguments.Config.FormPane).Visibility; Ring = [string]$w.FindName($Arguments.Config.RingPane).Visibility }
            }
            finally { if ($w) { $w.Close() } }
        }
        $r.Form | Should -Be 'Visible'
        $r.Ring | Should -Be 'Collapsed'
    }
}

Describe 'Dialog windows' -Skip:(-not $script:WindowGuard.Allowed) {
    It 'finds the dialog XAML in the script' {
        $script:DialogXaml.Count | Should -BeGreaterThan 0
    }

    It 'parses dialog <Index> as a MetroWindow with named controls' -ForEach $script:DialogCases {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:DialogXaml[$Index] } -Script {
            [xml]$x = $Arguments.Xaml
            $w = $null
            try {
                $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
                # Names inside a ControlTemplate belong to the template, not to the window, so FindName does not see them.
                $names = @($x.SelectNodes('//*[@*[local-name()="Name"]]') | Where-Object { -not $_.SelectSingleNode('ancestor::*[local-name()="ControlTemplate"]') } |
                        ForEach-Object { $_.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml') } | Where-Object { $_ })
                [PSCustomObject]@{ Type = $w.GetType().Name; Missing = @($names | Where-Object { -not $w.FindName($_) }); Count = @($names).Count }
            }
            finally { if ($w) { $w.Close() } }
        }
        $r.Type | Should -Be 'MetroWindow'
        $r.Missing | Should -BeNullOrEmpty
        $r.Count | Should -BeGreaterThan 0
    }
}

Describe 'Layout <Theme> <Size>' -ForEach $script:Matrix -Skip:(-not $script:WindowGuard.Allowed) {
    It 'has no clipped text in the <View> view' -ForEach $script:ViewCases {
        $m = ($script:Measured['{0}|{1}' -f $Theme, $Size] | ForEach-Object { $_ } | Where-Object View -eq $View)
        $m | Should -Not -BeNullOrEmpty -Because 'the view was measured'
        $m.Clipped | ForEach-Object { '{0}: needs {1}px, has {2}px' -f $_.Text, $_.Needed, $_.Actual } | Should -BeNullOrEmpty
    }

    It 'keeps every button, box and list inside the window in the <View> view' -ForEach $script:ViewCases {
        $m = ($script:Measured['{0}|{1}' -f $Theme, $Size] | ForEach-Object { $_ } | Where-Object View -eq $View)
        $m | Should -Not -BeNullOrEmpty -Because 'the view was measured'
        # Known defect: at 1024x720 the ring view needs more height than the window has, so the log box starts below the window.
        # The test lists it, so a new defect fails and a fixed one asks for the entry to go. The fix belongs to the layout of the ring view.
        $known = @('Ring Deployment|1024x720|txtLog')
        $found = @($m.Offscreen | ForEach-Object { '{0}|{1}|{2}' -f $View, $Size, $_.Name })
        $new = @($found | Where-Object { $known -notcontains $_ })
        $fixed = @($known | Where-Object { $_ -like ('{0}|{1}|*' -f $View, $Size) -and $found -notcontains $_ })
        $detail = @($m.Offscreen | ForEach-Object { '{0} at {1},{2} size {3}x{4}' -f $_.Name, $_.X, $_.Y, $_.W, $_.H })
        $new | Should -BeNullOrEmpty -Because ('a control is outside the window: ' + ($detail -join '; '))
        $fixed | Should -BeNullOrEmpty -Because ('a known defect is fixed: remove it from the list of this test: ' + ($fixed -join ', '))
    }
}
