#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Window test for the chrome and the buttons of the main window and of every dialog, in both themes.
# Every dialog of the app is measured:
#   - the dialogs that the start script builds from XAML text (here-strings in the functions Show-...Dialog) are loaded,
#     given the theme the way the script gives it (Set-DialogTheme) and shown. The Options dialog builds its panels in
#     code, and every panel button takes a style by SetResourceReference, so one sample button of each of those styles
#     is added to the content area and measured with the dialog.
#   - the dialogs that the start script takes from SuiteCommon (Show-ThemedMessage in its OK, OK and Cancel, and Yes and
#     No forms, and Show-CollectionPickerDialog) are opened by the real function with the main window as owner, and
#     measured and closed by a timer while they are open
#   - the main window is measured in the deployment view and in the ring view
# Rules:
#   - the dialog title bar brush (active and inactive) is the title bar brush of the main window
#   - the title bar text of the dialog is readable (4.5:1) on that brush
#   - every button has bold text (SemiBold or heavier, 600; a normal weight is a defect)
#   - no button has a border thicker than 1 px, and no button border has the text color of the theme

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
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:RepoDir = Split-Path -Parent $PSScriptRoot
    $shellPath = Join-Path $script:RepoDir $script:AppConfig.Shell
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($shellPath, [ref]$null, [ref]$null)

    # A string that holds the XAML of a dialog: a constant here-string or an expandable one ($titleText, $okText).
    $isDialogString = {
        param($n)
        ($n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) -and $n.Value -match '^\s*<Controls:MetroWindow'
    }
    $script:OwnDialogFunction = @()
    $dialogXaml = @()
    foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        foreach ($s in @($fn.Body.FindAll($isDialogString, $true))) {
            $script:OwnDialogFunction += $fn.Name
            $dialogXaml += $s.Value.Replace($script:SharedStyleMarker, (Get-SharedStyleXaml))
        }
    }

    # The dialogs of SuiteCommon that the start script calls.
    $suiteAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoDir 'Lib\SuiteCommon\SuiteCommon.psm1'), [ref]$null, [ref]$null)
    $suiteDialog = @($suiteAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and @($n.Body.FindAll($isDialogString, $true)).Count -gt 0 }, $true) | ForEach-Object { $_.Name })
    $called = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Sort-Object -Unique)
    $script:SharedDialogFunction = @($suiteDialog | Where-Object { $called -contains $_ } | Sort-Object -Unique)
    $shellText = Get-Content -LiteralPath $shellPath -Raw
    # One XamlReader load for the main window, one for each dialog.
    $script:ShellDialogLoads = @([regex]::Matches($shellText, 'XamlReader\]::Load')).Count - 1
    # The styles that the panel buttons of the Options dialog take: SetResourceReference(StyleProperty, 'Key').
    $panelStyle = @([regex]::Matches($shellText, "SetResourceReference\(\[System\.Windows\.Controls\.Control\]::StyleProperty, '(\w+)'\)") | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    $script:PanelStyle = $panelStyle

    $script:Measured = @{}
    foreach ($dark in $true, $false) {
        $key = if ($dark) { 'dark' } else { 'light' }
        $script:Measured[$key] = Invoke-WpfSta -Arguments @{ Dark = $dark; Config = $script:AppConfig; Dialogs = $dialogXaml; Shared = $script:SharedDialogFunction; PanelStyle = $panelStyle } -Script {
            $cfg = $Arguments.Config
            $main = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark
            Initialize-ShellTheme -Window $main -Dark $Arguments.Dark -WorkflowButtonNames $cfg.WorkflowButtons -OptionsButtonNames $cfg.OptionsButtons -LogLabelName $cfg.LogLabel
            $main.FindName($cfg.FormPane).Visibility = 'Visible'; $main.FindName($cfg.RingPane).Visibility = 'Collapsed'
            Wait-Layout $main
            $mainChrome = Get-SuiteWindowChrome -Window $main -Label 'main'
            $main.FindName($cfg.FormPane).Visibility = 'Collapsed'; $main.FindName($cfg.RingPane).Visibility = 'Visible'
            Wait-Layout $main
            $ringChrome = Get-SuiteWindowChrome -Window $main -Label 'main rings'
            $main.FindName($cfg.FormPane).Visibility = 'Visible'; $main.FindName($cfg.RingPane).Visibility = 'Collapsed'
            Wait-Layout $main
            $out = [pscustomobject]@{
                ThemeText = $mainChrome.ThemeText
                MainTitle = $mainChrome.Active; MainTitleInactive = $mainChrome.Inactive
                Main = @($mainChrome.Buttons) + @($ringChrome.Buttons); Dialogs = @(); Unmeasured = @()
            }
            $left = if ($env:SUITE_VISIBLE_WINDOWS -eq '1') { 120 } else { -32000 }
            $i = 0
            foreach ($xaml in $Arguments.Dialogs) {
                $i++
                [xml]$x = $xaml
                $d = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
                $d.WindowStartupLocation = 'Manual'
                $d.Left = $left; $d.Top = $left
                $d.ShowInTaskbar = $false
                Set-DialogTheme -Dialog $d
                $content = $d.FindName('contentArea')
                if ($content) {
                    # The Options dialog fills contentArea with panels at run time: one sample button for each panel button style.
                    $sample = New-Object System.Windows.Controls.StackPanel
                    foreach ($key in $Arguments.PanelStyle) {
                        $b = New-Object System.Windows.Controls.Button
                        $b.Name = 'sample' + $key; $b.Content = 'Panel ' + $key
                        [MahApps.Metro.Controls.ControlsHelper]::SetContentCharacterCasing($b, [System.Windows.Controls.CharacterCasing]::Normal)
                        $b.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, $key)
                        [void]$sample.Children.Add($b)
                    }
                    $content.Content = $sample
                }
                $d.Show(); Wait-Layout $d
                $c = Get-SuiteWindowChrome -Window $d -Label ([string]$d.Title)
                $out.Dialogs += [pscustomobject]@{ Title = $c.Title; Index = $i; Source = 'script'; Variant = ''; Active = $c.Active; Inactive = $c.Inactive; TitleRatio = $c.TitleRatio; Buttons = @($c.Buttons) }
                $d.Close()
            }

            # SuiteCommon dialogs: the real function opens the dialog modal; a timer measures and closes it.
            $collections = @(
                [pscustomobject]@{ CollectionID = 'MCM00001'; Name = 'All Systems'; FolderID = 0; MemberCount = 3; IsBuiltIn = $false }
                [pscustomobject]@{ CollectionID = 'MCM00002'; Name = 'Servers'; FolderID = 1; MemberCount = 2; IsBuiltIn = $false }
            )
            $folders = @([pscustomobject]@{ FolderID = 1; ParentID = 0; Name = 'Deployments' })
            $recipes = @()
            foreach ($shared in $Arguments.Shared) {
                if ($shared -eq 'Show-ThemedMessage') { foreach ($v in 'OK', 'OKCancel', 'YesNo') { $recipes += @{ Name = $shared; Variant = $v } } }
                else { $recipes += @{ Name = $shared; Variant = '' } }
            }
            foreach ($recipe in $recipes) {
                $name = $recipe.Name; $variant = $recipe.Variant
                $holder = @{ Chrome = $null; Error = $null }
                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(700)
                $mainWindow = $main
                $timer.Add_Tick({
                        $timer.Stop()
                        try {
                            $d = @($mainWindow.OwnedWindows) | Where-Object { $_ -is [MahApps.Metro.Controls.MetroWindow] } | Select-Object -First 1
                            Wait-Layout $d
                            $holder.Chrome = Get-SuiteWindowChrome -Window $d -Label ([string]$d.Title)
                            $d.Close()
                        }
                        catch { $holder.Error = $_.ToString(); foreach ($o in @($mainWindow.OwnedWindows)) { $o.Close() } }
                    }.GetNewClosure())
                $watchdog = New-Object System.Windows.Threading.DispatcherTimer
                $watchdog.Interval = [TimeSpan]::FromSeconds(10)
                $watchdog.Add_Tick({ $holder.Error = "watchdog closed the dialog"; foreach ($o in @($mainWindow.OwnedWindows)) { $o.Close() }
                    }.GetNewClosure())
                $timer.Start(); $watchdog.Start()
                try {
                    switch ($name) {
                        'Show-ThemedMessage' { [void](Show-ThemedMessage -Owner $main -Title 'Confirm Deployment' -Message 'Create the deployment?' -Buttons $variant -Icon Question) }
                        'Show-CollectionPickerDialog' { [void](Show-CollectionPickerDialog -Owner $main -Collections $collections -Folders $folders -Title 'Pick Device Collection') }
                        default { $timer.Stop(); $holder.Error = 'no recipe for ' + $name }
                    }
                }
                finally { $timer.Stop(); $watchdog.Stop() }
                if ($holder.Chrome) {
                    $c = $holder.Chrome
                    $out.Dialogs += [pscustomobject]@{ Title = $c.Title; Index = 0; Source = $name; Variant = $variant; Active = $c.Active; Inactive = $c.Inactive; TitleRatio = $c.TitleRatio; Buttons = @($c.Buttons) }
                }
                else { $out.Unmeasured += ('{0} {1}: {2}' -f $name, $variant, $holder.Error) }
            }
            $main.Close()
            $out
        }
    }
}

AfterAll {
    if ($script:WindowGuard.Allowed) { Close-WpfSession }
}

Describe 'Chrome <Theme>' -ForEach $script:ThemeCases -Skip:(-not $script:WindowGuard.Allowed) {
    BeforeAll { $script:M = $script:Measured[$Theme] }

    It 'finds every dialog of the app: the dialogs of the script and the dialogs of SuiteCommon that it calls' {
        # One dialog per function Show-...Dialog of the script that builds a window.
        $funcs = @([regex]::Matches((Get-Content -LiteralPath (Join-Path $script:RepoDir $script:AppConfig.Shell) -Raw), '(?m)^function Show-\w*Dialog\b'))
        $funcs.Count | Should -BeGreaterThan 3
        @($script:OwnDialogFunction | Sort-Object -Unique).Count | Should -Be $funcs.Count
        $script:ShellDialogLoads | Should -Be $funcs.Count
        @($script:M.Dialogs | Where-Object { $_.Source -eq 'script' }).Count | Should -Be $funcs.Count
        $script:SharedDialogFunction | Should -Contain 'Show-ThemedMessage'
        $script:SharedDialogFunction | Should -Contain 'Show-CollectionPickerDialog'
        @($script:M.Dialogs | Where-Object { $_.Source -ne 'script' } | ForEach-Object Source | Sort-Object -Unique).Count | Should -Be $script:SharedDialogFunction.Count
        @($script:M.Dialogs | Where-Object { $_.Source -eq 'Show-ThemedMessage' }).Count | Should -Be 3
        @($script:M.Unmeasured) | Should -BeNullOrEmpty
    }

    It 'measures a sample button of every style that the Options panels take' {
        $script:PanelStyle | Should -Contain 'DialogButton'
        $script:PanelStyle | Should -Contain 'DialogAccentButton'
        $options = @($script:M.Dialogs | Where-Object { $_.Title -eq 'Options' })
        $options.Count | Should -Be 1
        foreach ($key in $script:PanelStyle) { @($options[0].Buttons | Where-Object { $_.Name -eq ('sample' + $key) }).Count | Should -Be 1 }
    }

    It 'gives every dialog the title bar brush of the main window' {
        $bad = @($script:M.Dialogs | Where-Object { $_.Active -ne $script:M.MainTitle -or $_.Inactive -ne $script:M.MainTitleInactive } |
                ForEach-Object { '{0}: active {1} inactive {2}, main {3} {4}' -f $_.Title, $_.Active, $_.Inactive, $script:M.MainTitle, $script:M.MainTitleInactive })
        $bad | Should -BeNullOrEmpty
    }

    It 'keeps the title text of every dialog readable at 4.5:1' {
        $bad = @($script:M.Dialogs | Where-Object { $_.TitleRatio -lt 4.5 } | ForEach-Object { '{0}: {1}' -f $_.Title, $_.TitleRatio })
        $bad | Should -BeNullOrEmpty
    }

    It 'shows bold text on every button of the main window and of every dialog' {
        $all = @($script:M.Main) + @($script:M.Dialogs | ForEach-Object { $_.Buttons })
        $all.Count | Should -BeGreaterThan 5
        $bad = @($all | Where-Object { $_.Weight -lt 600 } | ForEach-Object { '{0} "{1}" weight {2}' -f $_.Window, $_.Text, $_.Weight })
        $bad | Should -BeNullOrEmpty
    }

    It 'gives no button a border thicker than 1 px or in the text color of the theme' {
        $all = @($script:M.Main) + @($script:M.Dialogs | ForEach-Object { $_.Buttons })
        $bad = @($all | Where-Object { $_.BorderThickness -gt 1 -or ($_.BorderThickness -gt 0 -and $_.BorderColor -eq $script:M.ThemeText) } |
                ForEach-Object { '{0} "{1}" border {2}px {3}' -f $_.Window, $_.Text, $_.BorderThickness, $_.BorderColor })
        $bad | Should -BeNullOrEmpty
    }
}
