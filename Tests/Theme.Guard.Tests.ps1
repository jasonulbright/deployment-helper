#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard tests for the theme rules. They parse the files, so they run on Windows PowerShell 5.1 and
# PowerShell 7. One test creates a real button in its own runspace to check the foreground helper.
#   (a) Every blue background in XAML and PowerShell has a white foreground.
#   (b) No font size, weight, style or text decoration depends on the theme.
#   (c) No banned word is in a tooltip, a UI string, the README or the user guide.
#   (d) The site code and SMS Provider titles are bold 14 pt, the values are not.
#   (e) Every button takes the shared themed style, and no style holds a color literal.
#   (f) No raw exception text reaches the UI.

BeforeDiscovery {
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:PsFilesD = @(
        'start-deploymenthelper.ps1', 'Module\DeploymentHelperCommon.psm1', 'Lib\SuiteCommon\SuiteCommon.psm1'
    ) | ForEach-Object { @{ Name = $_; Path = (Join-Path $script:RepoRoot $_) } }
    $script:AppPsFilesD = @(
        'start-deploymenthelper.ps1', 'Module\DeploymentHelperCommon.psm1'
    ) | ForEach-Object { @{ Name = $_; Path = (Join-Path $script:RepoRoot $_) } }
    $script:XamlDocsD = @(Get-XamlDocument)
}

BeforeAll {
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:XamlDocs = @(Get-XamlDocument)
    $script:ShellName = 'start-deploymenthelper.ps1'
    $script:ShellPath = Join-Path $script:RepoRoot $script:ShellName
    $script:ModulePath = Join-Path $script:RepoRoot 'Module\DeploymentHelperCommon.psm1'
}

Describe 'Blue backgrounds have white text (a)' {
    It 'finds the XAML of the main window and of the four dialogs' {
        $script:XamlDocs.Count | Should -BeGreaterOrEqual 5
        @($script:XamlDocs | Where-Object { $_.Source -eq 'MainWindow.xaml' }).Count | Should -Be 1
        @($script:XamlDocs | Where-Object { $_.Source -like "$($script:ShellName) here-string*" }).Count | Should -BeGreaterOrEqual 4
    }

    It 'sets a white foreground on every blue background in <_.Source>' -ForEach $script:XamlDocsD {
        $bad = @(Get-BlueWithoutWhite -Text $_.Text -Source $_.Source)
        $bad | Should -BeNullOrEmpty
    }

    It 'recognizes the system highlight color as blue' {
        Test-BlueBrush '{DynamicResource {x:Static SystemColors.HighlightBrushKey}}' | Should -BeTrue
        Test-BlueBrush 'SystemColors.HighlightColor' | Should -BeTrue
    }

    It 'flags a blue nested Background property element that has no white Foreground' {
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
        $bad = "<Border $ns><Border.Background><SolidColorBrush Color=`"{DynamicResource MahApps.Colors.AccentBase}`"/></Border.Background></Border>"
        $ok = "<Border $ns TextElement.Foreground=`"White`"><Border.Background><SolidColorBrush Color=`"#0078D4`"/></Border.Background></Border>"
        @(Get-BlueWithoutWhite -Text $bad -Source 'sample').Count | Should -Be 1
        @(Get-BlueWithoutWhite -Text $ok -Source 'sample').Count | Should -Be 0
    }

    It 'recognizes blue and white brushes' {
        Test-BlueBrush '{DynamicResource MahApps.Brushes.Accent}' | Should -BeTrue
        Test-BlueBrush '{DynamicResource MahApps.Colors.Accent}' | Should -BeTrue
        Test-BlueBrush '{DynamicResource MahApps.Colors.AccentBase}' | Should -BeTrue
        Test-BlueBrush '#0078D4' | Should -BeTrue
        Test-BlueBrush '#1E1E1E' | Should -BeFalse
        Test-BlueBrush '#3A3A3A' | Should -BeFalse
        Test-WhiteBrush '{DynamicResource MahApps.Brushes.IdealForeground}' | Should -BeTrue
        Test-WhiteBrush 'Black' | Should -BeFalse
    }

    It 'flags a blue background that has no white foreground, also with a qualified property and a nested color' {
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
        $setter = "<Style $ns TargetType=`"Button`"><Setter Property=`"Background`" Value=`"#0078D4`"/></Style>"
        @(Get-BlueWithoutWhite -Text $setter -Source 'sample').Count | Should -Be 1
        $ok = "<Style $ns TargetType=`"Button`"><Setter Property=`"Background`" Value=`"#0078D4`"/><Setter Property=`"Foreground`" Value=`"White`"/></Style>"
        @(Get-BlueWithoutWhite -Text $ok -Source 'sample').Count | Should -Be 0
        $attr = "<Button $ns Background=`"{DynamicResource MahApps.Brushes.Accent}`"/>"
        @(Get-BlueWithoutWhite -Text $attr -Source 'sample').Count | Should -Be 1
        $qualified = "<ControlTemplate $ns TargetType=`"Button`"><ControlTemplate.Triggers><Trigger Property=`"IsPressed`" Value=`"True`"><Setter Property=`"Border.Background`" Value=`"{DynamicResource MahApps.Brushes.Accent}`"/></Trigger></ControlTemplate.Triggers></ControlTemplate>"
        @(Get-BlueWithoutWhite -Text $qualified -Source 'sample').Count | Should -Be 1
        $nested = "<Style $ns TargetType=`"Button`"><Setter Property=`"Background`"><Setter.Value><SolidColorBrush Color=`"{DynamicResource MahApps.Colors.Accent}`"/></Setter.Value></Setter></Style>"
        @(Get-BlueWithoutWhite -Text $nested -Source 'sample').Count | Should -Be 1
    }

    It 'sets white text on the selected state of <Target> in the main window' -ForEach @(
        @{ Target = 'GridRowStyle'; Kind = 'key' }
        @{ Target = 'GridCellStyle'; Kind = 'key' }
        @{ Target = 'ComboBoxItem'; Kind = 'type' }
        @{ Target = 'ListBoxItem'; Kind = 'type' }
    ) {
        $doc = @($script:XamlDocs | Where-Object { $_.Source -eq 'MainWindow.xaml' })[0]
        [xml]$x = $doc.Text
        $styles = @($x.SelectNodes("//*[local-name()='Style']") | Where-Object {
                if ($Kind -eq 'key') { $_.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml') -eq $Target }
                else { $_.GetAttribute('TargetType') -eq $Target -and -not $_.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml') }
            })
        $styles.Count | Should -BeGreaterThan 0
        foreach ($st in $styles) {
            $trigger = @($st.SelectNodes("*[local-name()='Style.Triggers']/*[local-name()='Trigger']") | Where-Object { $_.GetAttribute('Property') -in 'IsSelected', 'IsHighlighted' })
            $trigger.Count | Should -BeGreaterThan 0
            if ($Target -eq 'ComboBoxItem') { @($trigger | ForEach-Object { $_.GetAttribute('Property') } | Sort-Object) | Should -Be @('IsHighlighted', 'IsSelected') }
            foreach ($tr in $trigger) {
                $fg = @($tr.SelectNodes("*[local-name()='Setter']") | Where-Object { $_.GetAttribute('Property') -eq 'Foreground' })
                $fg.Count | Should -Be 1
                Test-WhiteBrush (Get-SetterValue $fg[0]) | Should -BeTrue
            }
        }
    }

    It 'gives every DataGrid of the app the shared row and cell styles' {
        $grids = 0
        foreach ($d in $script:XamlDocs) {
            [xml]$x = $d.Text
            foreach ($g in @($x.SelectNodes("//*[local-name()='DataGrid']"))) {
                $grids++
                $g.GetAttribute('RowStyle') | Should -Match 'GridRowStyle' -Because ('{0} {1}' -f $d.Source, $g.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml'))
                $g.GetAttribute('CellStyle') | Should -Match 'GridCellStyle' -Because ('{0} {1}' -f $d.Source, $g.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml'))
            }
        }
        $grids | Should -BeGreaterOrEqual 3
    }

    It 'gives each dialog with a list or a combo box the shared item styles with white selected text' {
        $dialogs = @($script:XamlDocs | Where-Object { $_.Source -like "$($script:ShellName) here-string*" -and $_.Text -match 'MetroWindow' })
        $dialogs.Count | Should -BeGreaterOrEqual 4
        $listStyles = 0
        foreach ($d in $dialogs) {
            [xml]$x = $d.Text
            foreach ($type in 'ComboBoxItem', 'ListBoxItem') {
                $uses = ($type -eq 'ComboBoxItem' -and $x.SelectNodes("//*[local-name()='ComboBox']").Count -gt 0) -or ($type -eq 'ListBoxItem' -and $x.SelectNodes("//*[local-name()='ListBox']").Count -gt 0)
                if (-not $uses) { continue }
                $st = @($x.SelectNodes("//*[local-name()='Style']") | Where-Object { $_.GetAttribute('TargetType') -eq $type -and -not $_.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml') })
                $st.Count | Should -BeGreaterOrEqual 1 -Because ('{0} needs a {1} style' -f $d.Source, $type)
                if ($type -eq 'ListBoxItem') { $listStyles++ }
                $setters = @($st[0].SelectNodes('.//*[local-name()="Trigger"]/*[local-name()="Setter"]') | Where-Object { $_.GetAttribute('Property') -eq 'Foreground' })
                $setters.Count | Should -BeGreaterThan 0
                foreach ($se in $setters) { Test-WhiteBrush (Get-SetterValue $se) | Should -BeTrue }
            }
        }
        $listStyles | Should -BeGreaterOrEqual 2
    }

    It 'has no blue brush literal in <Name>, and each CSS rule with a blue background has white text' -ForEach $script:AppPsFilesD {
        $bad = foreach ($s in @(Get-ScriptString -Path $Path)) {
            if ($s -notmatch '<' -and (Test-BlueBrush $s)) { 'literal ' + $s }
            foreach ($m in [regex]::Matches($s, '\{([^}]*)\}')) {
                $rule = $m.Groups[1].Value
                $bg = [regex]::Match($rule, '(?i)(?<![-\w])background(?:-color)?\s*:\s*(#[0-9A-F]{6,8})\b')
                if ($bg.Success -and (Test-BlueBrush $bg.Groups[1].Value) -and $rule -notmatch '(?i)(?<![-\w])color\s*:\s*(#fff(fff)?|white)\b') { 'css ' + $rule }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'sets the title bar brushes through SuiteCommon and not with a color of its own' {
        $text = Get-Content -LiteralPath $script:ShellPath -Raw
        $text | Should -Match 'Update-TitleBarBrushes\s+-IsDark\s+\$true'
        $text | Should -Match 'Update-TitleBarBrushes\s+-IsDark\s+\$false'
        $text | Should -Not -Match 'WindowTitleBrush\s*='
    }

    It 'paints the light theme buttons of SuiteCommon with white text' {
        $path = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match "OnBlue\s*=\s*\`$bc\.ConvertFrom\('#FFFFFF'\)"
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
        foreach ($name in 'Update-SidebarButtonTheme', 'Set-ButtonTheme') {
            $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
            $fn | Should -Not -BeNullOrEmpty -Because $name
            $fn.Extent.Text | Should -Match 'Set-ThemeForeground' -Because ('{0} sets the foreground with each background' -f $name)
        }
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Set-ThemeForeground' }, $true)
        $fn.Extent.Text | Should -Match '\.Foreground\s*=\s*\$OnBlue'
    }
}

Describe 'Fonts do not depend on the theme (b)' {
    It 'has no font property in a theme branch of <Name>' -ForEach $script:PsFilesD {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
        $assigns = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] }, $true)
        $bad = foreach ($a in $assigns) {
            $member = [string]$a.Left.Member.Value
            if ($script:ThemeProperty -notcontains $member) { continue }
            $p = $a.Parent
            while ($p) {
                $condText = ''
                if ($p -is [System.Management.Automation.Language.IfStatementAst]) { $condText = (@($p.Clauses | ForEach-Object { $_.Item1.Extent.Text }) -join ' ') }
                elseif ($p -is [System.Management.Automation.Language.SwitchStatementAst]) { $condText = $p.Condition.Extent.Text }
                if ($condText -match '(?i)IsDark|DarkMode|Light|Theme|IsOn') { '{0}:{1} sets {2} in a theme branch' -f $Name, $a.Extent.StartLineNumber, $member; break }
                $p = $p.Parent
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'detects a font property in a theme branch (the check itself)' {
        $ast = [System.Management.Automation.Language.Parser]::ParseInput('if ($isDark) { $b.Background = 1; $b.FontWeight = 2 } else { $b.Background = 3 }', [ref]$null, [ref]$null)
        $assigns = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and $script:ThemeProperty -contains [string]$n.Left.Member.Value }, $true))
        $assigns.Count | Should -Be 1
        $assigns[0].Parent.Parent | Should -BeOfType [System.Management.Automation.Language.IfStatementAst]
    }

    It 'sets no font property in the theme functions of SuiteCommon except the connection boxes' {
        $path = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
        foreach ($name in 'Set-ThemeForeground', 'Update-SidebarButtonTheme', 'Set-ButtonTheme', 'Set-DialogTheme', 'Update-TitleBarBrushes', 'Initialize-SuiteTheme') {
            $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
            $fn | Should -Not -BeNullOrEmpty -Because $name
            $fn.Extent.Text | Should -Not -Match '\.(FontSize|FontWeight|FontStyle|TextDecorations)\s*=' -Because $name
        }
    }

    It 'makes the field titles bold and 14 pt or larger the same way in both themes' {
        $path = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match 'FieldTitleFontSize\s*=\s*14'
        $text | Should -Match '\.FontWeight\s*=\s*\[System\.Windows\.FontWeights\]::Bold'
    }

    It 'uses no theme trigger on a font property in <_.Source>' -ForEach $script:XamlDocsD {
        @(Get-ThemeFontTrigger -Text $_.Text) | Should -BeNullOrEmpty
    }

    It 'flags a trigger that sets a font for a theme, and not one that sets it for the selected state (the check itself)' {
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
        $theme = "<Style $ns TargetType=`"TextBlock`"><Style.Triggers><DataTrigger Binding=`"{Binding IsDark}`" Value=`"True`"><Setter Property=`"FontWeight`" Value=`"Bold`"/></DataTrigger></Style.Triggers></Style>"
        $state = "<Style $ns TargetType=`"TextBlock`"><Style.Triggers><DataTrigger Binding=`"{Binding IsSelected, RelativeSource={RelativeSource AncestorType=TabItem}}`" Value=`"True`"><Setter Property=`"FontWeight`" Value=`"Bold`"/></DataTrigger></Style.Triggers></Style>"
        @(Get-ThemeFontTrigger -Text $theme).Count | Should -Be 1
        @(Get-ThemeFontTrigger -Text $state).Count | Should -Be 0
    }
}

Describe 'No banned word in the UI text and the documents (c)' {
    It 'finds the banned word <Word>' -ForEach @(@{ Word = 'simply' }, @{ Word = 'Helps you' }, @{ Word = 'streamlines' }) {
        $Word | Should -Match $script:BannedPattern
    }

    It 'has no banned word in the XAML text of <_.Source>' -ForEach $script:XamlDocsD {
        $bad = @(Get-XamlUiText -Text $_.Text | Where-Object { $_ -match $script:BannedPattern })
        $bad | Should -BeNullOrEmpty
    }

    It 'has no banned word in the strings of <Name>' -ForEach $script:AppPsFilesD {
        $bad = @(Get-ScriptString -Path $Path | Where-Object { $_ -match $script:BannedPattern })
        $bad | Should -BeNullOrEmpty
    }

    It 'has no banned word in the README or the user guide' {
        $files = @(Get-Item -LiteralPath (Join-Path $script:RepoRoot 'README.md')) + @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'docs') -Filter '*.md' -File -Recurse)
        $files.Count | Should -BeGreaterThan 5
        $bad = foreach ($f in $files) {
            $n = 0
            foreach ($line in (Get-Content -LiteralPath $f.FullName)) {
                $n++
                if ($line -match $script:BannedPattern) { '{0}:{1}: {2}' -f $f.Name, $n, $Matches[0] }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'gives every tooltip in <_.Source> one short sentence' -ForEach $script:XamlDocsD {
        $bad = @(Get-XamlToolTip -Text $_.Text | Where-Object { -not (Test-OneSentence $_) })
        $bad | Should -BeNullOrEmpty
    }

    It 'checks one sentence (the check itself)' {
        Test-OneSentence 'Deploy an application to a device collection.' | Should -BeTrue
        Test-OneSentence 'Pick the device collection (Shift+click reloads the list).' | Should -BeTrue
        Test-OneSentence 'Pick the object. Then select OK.' | Should -BeFalse
        Test-OneSentence 'pick the object.' | Should -BeFalse
        Test-OneSentence 'Pick the object' | Should -BeFalse
    }

    It 'uses Title Case for the text of every button' {
        $bad = @()
        $count = 0
        foreach ($d in $script:XamlDocs) {
            [xml]$x = $d.Text
            foreach ($b in @($x.SelectNodes("//*[local-name()='Button'][@Content]"))) {
                $count++
                if (-not (Test-TitleCase $b.GetAttribute('Content'))) { $bad += ('{0}: {1}' -f $d.Source, $b.GetAttribute('Content')) }
            }
        }
        foreach ($b in @(Get-CodeControl -Path $script:ShellPath -TypeName 'Button')) {
            $content = (Get-MemberAssignment -Node $b.Function -Variable $b.Variable)['Content']
            if ($content -match "^'([^']*)'$") {
                $count++
                if (-not (Test-TitleCase $Matches[1])) { $bad += ('{0}: {1}' -f $b.Variable, $Matches[1]) }
            }
        }
        $count | Should -BeGreaterThan 15
        @($bad) | Should -BeNullOrEmpty
    }

    It 'checks Title Case (the check itself)' {
        Test-TitleCase 'Refresh and Reconcile' | Should -BeTrue
        Test-TitleCase 'Select DP Groups...' | Should -BeTrue
        Test-TitleCase 'Select DP groups...' | Should -BeFalse
        Test-TitleCase 'Save template' | Should -BeFalse
    }
}

Describe 'Theme helpers of SuiteCommon behave and are used (a, b, d)' {
    BeforeAll {
        $script:SuitePath = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $script:SuiteAst = [System.Management.Automation.Language.Parser]::ParseFile($script:SuitePath, [ref]$null, [ref]$null)
        function Get-SuiteFunction {
            param([string]$Name)
            $script:SuiteAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
        }
        function Get-CommandCount {
            param($Ast, [string]$Name)
            @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq $Name }, $true)).Count
        }
        # The test loads the WPF assemblies itself, so that it does not depend on an earlier test or on the shell.
        Add-Type -AssemblyName PresentationFramework, WindowsBase
        . ([scriptblock]::Create((Get-SuiteFunction 'Set-ThemeForeground').Extent.Text))
    }

    It 'Set-ThemeForeground sets the on-blue brush in light mode and clears it in dark mode' {
        $c = [pscustomobject]@{ Foreground = $null }
        Set-ThemeForeground -Control $c -IsDark $false -OnBlue 'White'
        $c.Foreground | Should -Be 'White'
        Set-ThemeForeground -Control $c -IsDark $true -OnBlue 'White'
        $c.Foreground | Should -BeNullOrEmpty
    }

    It 'Set-ThemeForeground sets and clears the foreground of a real Button (STA runspace)' {
        $fnText = (Get-SuiteFunction 'Set-ThemeForeground').Extent.Text
        $rs = [runspacefactory]::CreateRunspace()
        $rs.ApartmentState = [System.Threading.ApartmentState]::STA
        $rs.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $rs.Open()
        try {
            $ps = [powershell]::Create()
            $ps.Runspace = $rs
            [void]$ps.AddScript({
                    param($fnText)
                    Add-Type -AssemblyName PresentationFramework, WindowsBase
                    . ([scriptblock]::Create($fnText))
                    $button = New-Object System.Windows.Controls.Button
                    $white = [System.Windows.Media.Brushes]::White
                    Set-ThemeForeground -Control $button -IsDark $false -OnBlue $white
                    $light = [object]::ReferenceEquals($button.Foreground, $white)
                    Set-ThemeForeground -Control $button -IsDark $true -OnBlue $white
                    $dark = [object]::ReferenceEquals($button.ReadLocalValue([System.Windows.Controls.Control]::ForegroundProperty), [System.Windows.DependencyProperty]::UnsetValue)
                    [pscustomobject]@{ Light = $light; Dark = $dark }
                }).AddArgument($fnText)
            $result = @($ps.Invoke())
            $ps.HadErrors | Should -BeFalse
            $result.Count | Should -Be 1
            $result[0].Light | Should -BeTrue
            $result[0].Dark | Should -BeTrue
        }
        finally {
            if ($ps) { $ps.Dispose() }
            $rs.Dispose()
        }
    }

    It '<Name> calls Set-ThemeForeground' -ForEach @(@{ Name = 'Set-ButtonTheme' }, @{ Name = 'Update-SidebarButtonTheme' }) {
        (Get-CommandCount (Get-SuiteFunction $Name) 'Set-ThemeForeground') | Should -BeGreaterThan 0
    }

    It '<Name> calls Set-ConnectionTextTheme' -ForEach @(@{ Name = 'Set-ButtonTheme' }, @{ Name = 'Update-SidebarButtonTheme' }) {
        (Get-CommandCount (Get-SuiteFunction $Name) 'Set-ConnectionTextTheme') | Should -BeGreaterThan 0
    }

    It 'the script starts the theme and every dialog calls Set-DialogTheme' {
        $shell = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
        (Get-CommandCount $shell 'Initialize-SuiteTheme') | Should -BeGreaterThan 0
        (Get-CommandCount $shell 'Set-ButtonTheme') | Should -BeGreaterOrEqual 3
        $dialogs = @($script:XamlDocs | Where-Object { $_.Source -like "$($script:ShellName) here-string*" -and $_.Text -match 'MetroWindow' }).Count
        $dialogs | Should -BeGreaterOrEqual 4
        (Get-CommandCount $shell 'Set-DialogTheme') | Should -BeGreaterOrEqual $dialogs
        (Get-SuiteFunction 'Set-DialogTheme') | Should -Not -BeNullOrEmpty
    }

    It 'sets a white foreground in every style based on an accent style, in <_.Source>' -ForEach $script:XamlDocsD {
        [xml]$x = $_.Text
        $bad = foreach ($st in @($x.SelectNodes("//*[local-name()='Style']"))) {
            # A style that is based on a MahApps accent style must set its own white foreground. A style that is based on DialogAccentButton inherits it.
            if ($st.GetAttribute('BasedOn') -match 'MahApps\.Styles\.[\w.]*Accent') {
                $fg = $st.ChildNodes | Where-Object { $_.LocalName -eq 'Setter' -and $_.GetAttribute('Property') -eq 'Foreground' } | Select-Object -First 1
                if (-not $fg -or -not (Test-WhiteBrush (Get-SetterValue $fg))) { 'Style {0} has no white Foreground' -f $st.GetAttribute('Key') }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'gives the site code and SMS Provider titles of the Options page bold 14 pt and the value boxes the body style' {
        $fn = Get-ScriptFunction -Path $script:ShellPath -Name 'New-ConnectionPanel'
        $fn | Should -Not -BeNullOrEmpty
        foreach ($pair in @(
                @{ Title = 'lblSite'; Box = 'txtSite'; Name = 'txtSiteCode' }
                @{ Title = 'lblProv'; Box = 'txtProv'; Name = 'txtSmsProvider' }
            )) {
            $title = Get-MemberAssignment -Node $fn -Variable $pair.Title
            $box = Get-MemberAssignment -Node $fn -Variable $pair.Box
            [double]$title['FontSize'] | Should -BeGreaterOrEqual 14 -Because $pair.Title
            $title['FontWeight'] | Should -Match 'Bold' -Because $pair.Title
            $box.ContainsKey('FontWeight') | Should -BeFalse -Because ('value box {0} keeps the normal weight' -f $pair.Box)
            [double]$box['FontSize'] | Should -BeLessThan 14 -Because ('value box {0}' -f $pair.Box)
            $box['Name'] | Should -Match $pair.Name
        }
    }

    It 'SuiteCommon bolds the titles with 14 pt and not the values' {
        $text = Get-Content -LiteralPath $script:SuitePath -Raw
        $text | Should -Match 'FieldTitleFontSize\s*=\s*14'
        $text | Should -Match 'Find-FieldTitle'
        $text | Should -Not -Match '\$box\.FontWeight'
    }

    It 'bolds every field label of the Options pages and keeps the page headers SemiBold' {
        $count = 0
        $headers = 0
        $bad = foreach ($fnName in 'New-ConnectionPanel', 'New-LoggingPanel', 'New-HistoryPanel', 'New-TemplatesPanel', 'New-AboutPanel') {
            $fn = Get-ScriptFunction -Path $script:ShellPath -Name $fnName
            $fn | Should -Not -BeNullOrEmpty -Because $fnName
            foreach ($tb in @(Get-CodeControl -Path $script:ShellPath -TypeName 'TextBlock' | Where-Object { $_.FunctionName -eq $fnName })) {
                $assigned = Get-MemberAssignment -Node $fn -Variable $tb.Variable
                if ($tb.Variable -match '^(hdr|dpHdr)$') {
                    $headers++
                    if ($assigned['FontWeight'] -notmatch 'SemiBold') { '{0}: ${1} is not SemiBold' -f $fnName, $tb.Variable }
                    continue
                }
                $isTitle = ($tb.Variable -match '^lbl' -and $tb.Variable -notmatch '^(lblStatus|lblDPStatus|lblHistoryInfo)$')
                if (-not $isTitle) { continue }
                $count++
                if ($assigned['FontWeight'] -notmatch 'Bold') { '{0}: ${1} is not bold' -f $fnName, $tb.Variable }
            }
        }
        $count | Should -BeGreaterOrEqual 10
        $headers | Should -BeGreaterOrEqual 6
        @($bad) | Should -BeNullOrEmpty
    }

    It 'bolds the field labels and the app title of the main window and keeps the headers SemiBold' {
        $doc = @($script:XamlDocs | Where-Object { $_.Source -eq 'MainWindow.xaml' })[0]
        [xml]$x = $doc.Text
        foreach ($pair in @(@{ Key = 'FormLabelStyle'; Weight = 'Bold' }, @{ Key = 'SectionHeaderStyle'; Weight = 'SemiBold' })) {
            $st = $x.SelectSingleNode(("//*[local-name()='Style'][@*[local-name()='Key']='{0}']" -f $pair.Key))
            $st | Should -Not -BeNullOrEmpty -Because $pair.Key
            $w = @($st.SelectNodes("*[local-name()='Setter'][@Property='FontWeight']"))
            $w.Count | Should -Be 1 -Because $pair.Key
            $w[0].GetAttribute('Value') | Should -Be $pair.Weight -Because $pair.Key
        }
        $labels = @($x.SelectNodes("//*[local-name()='TextBlock'][@Text][substring(@Text, string-length(@Text)) = ':']"))
        $labels.Count | Should -BeGreaterOrEqual 12
        $bad = foreach ($t in $labels) { if ($t.GetAttribute('Style') -notmatch 'FormLabelStyle') { 'Label {0} does not use FormLabelStyle' -f $t.GetAttribute('Text') } }
        @($bad) | Should -BeNullOrEmpty
        foreach ($pair in @(@{ Name = 'txtAppTitle'; Weight = 'Bold' }, @{ Name = 'txtModuleHeader'; Weight = 'SemiBold' }, @{ Name = 'lblLogOutput'; Weight = 'SemiBold' })) {
            $t = $x.SelectSingleNode(("//*[@*[local-name()='Name']='{0}']" -f $pair.Name))
            $t.GetAttribute('FontWeight') | Should -Be $pair.Weight -Because $pair.Name
        }
    }

    It 'gives the field values the normal weight in the main window' {
        $doc = @($script:XamlDocs | Where-Object { $_.Source -eq 'MainWindow.xaml' })[0]
        [xml]$x = $doc.Text
        $boxes = @($x.SelectNodes("//*[local-name()='TextBox' or local-name()='ComboBox' or local-name()='DateTimePicker']"))
        $boxes.Count | Should -BeGreaterThan 5
        foreach ($b in $boxes) { $b.GetAttribute('FontWeight') | Should -BeIn @('', 'Normal') -Because $b.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml') }
    }
}

Describe 'Every button takes the shared themed style (e)' {
    BeforeAll {
        $script:ShellText = Get-Content -LiteralPath $script:ShellPath -Raw
    }

    It 'gives each Button in <_.Source> a style that is based on a MahApps button style' -ForEach $script:XamlDocsD {
        @(Get-ButtonWithoutSharedStyle -Text $_.Text) | Should -BeNullOrEmpty
    }

    It 'detects a button without the shared style (the check itself)' {
        $head = '<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">'
        $bare = $head + '<Button x:Name="bare" Content="x"/></Grid>'
        $own = $head + '<Grid.Resources><Style x:Key="Own" TargetType="Button"/></Grid.Resources><Button x:Name="own" Style="{StaticResource Own}"/></Grid>'
        $good = $head + '<Grid.Resources><Style x:Key="Base" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}"/><Style x:Key="Inline" TargetType="Button" BasedOn="{StaticResource Base}"/></Grid.Resources><Button x:Name="good" Style="{StaticResource Inline}"/></Grid>'
        $nested = $head + '<Grid.Resources><Style x:Key="Base" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}"/></Grid.Resources><Button x:Name="nested"><Button.Style><Style TargetType="Button" BasedOn="{StaticResource Base}"/></Button.Style></Button></Grid>'
        @(Get-ButtonWithoutSharedStyle -Text $bare).Count | Should -Be 1
        @(Get-ButtonWithoutSharedStyle -Text $own).Count | Should -Be 1
        @(Get-ButtonWithoutSharedStyle -Text $good).Count | Should -Be 0
        @(Get-ButtonWithoutSharedStyle -Text $nested).Count | Should -Be 0
    }

    It 'gives every Button in the main window the themed style, and the sidebar buttons their own styles' {
        $doc = @($script:XamlDocs | Where-Object { $_.Source -eq 'MainWindow.xaml' })[0]
        [xml]$x = $doc.Text
        $buttons = @($x.SelectNodes("//*[local-name()='Button']"))
        $buttons.Count | Should -BeGreaterThan 12
        $wf = [regex]::Match($script:ShellText, '\$script:WorkflowButtons\s*=\s*@\(([^)]*)\)').Groups[1].Value
        $opt = [regex]::Match($script:ShellText, '\$script:OptionsButtons\s*=\s*@\(([^)]*)\)').Groups[1].Value
        $sidebar = @(([regex]::Matches(($wf + ' ' + $opt), '\$(\w+)') | ForEach-Object { $_.Groups[1].Value }))
        $sidebar.Count | Should -Be 6
        foreach ($b in $buttons) {
            $name = $b.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
            $chain = @(Get-ButtonStyleChain -Doc $x -Button $b)
            $chain | Should -Contain 'ThemedButtonStyle' -Because $name
            if ($sidebar -contains $name) { $chain | Should -Contain 'WorkflowButtonStyle' -Because $name }
            else { ($chain -contains 'DialogButton' -or $chain -contains 'DialogAccentButton') | Should -BeTrue -Because $name }
        }
        $script:ShellText | Should -Match '-WorkflowButtons\s+\$script:WorkflowButtons'
        $script:ShellText | Should -Match '-OptionsButtons\s+\$script:OptionsButtons'
    }

    It 'gives every dialog button the primary or the one secondary style' {
        $docs = @($script:XamlDocs | Where-Object { $_.Source -like "$($script:ShellName) here-string*" -and $_.Text -match 'MetroWindow' })
        $docs.Count | Should -BeGreaterOrEqual 4
        $bad = foreach ($d in $docs) {
            [xml]$x = $d.Text
            $buttons = @($x.SelectNodes("//*[local-name()='Button']"))
            $buttons.Count | Should -BeGreaterThan 0
            foreach ($b in $buttons) {
                $name = $b.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
                $chain = @(Get-ButtonStyleChain -Doc $x -Button $b)
                if (-not ($chain -contains 'DialogButton' -or $chain -contains 'DialogAccentButton')) { '{0}: {1} has neither the primary nor the secondary dialog style' -f $d.Source, $name }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'gives every button that the script creates one of the dialog styles' {
        $buttons = @(Get-CodeControl -Path $script:ShellPath -TypeName 'Button')
        $buttons.Count | Should -BeGreaterOrEqual 8
        $bad = foreach ($b in $buttons) {
            $key = Get-CodeStyleKey -Function $b.Function -Variable $b.Variable
            if ($key -notin 'DialogButton', 'DialogAccentButton') { 'line {0}: ${1} has the style "{2}"' -f $b.Line, $b.Variable, $key }
        }
        @($bad) | Should -BeNullOrEmpty
        foreach ($f in $script:ShellPath, $script:ModulePath) {
            (Get-Content -LiteralPath $f -Raw) | Should -Not -Match "MahApps\.Styles\.Button" -Because $f
        }
    }

    It 'detects a code-built button without a style (the check itself)' {
        $ast = [System.Management.Automation.Language.Parser]::ParseInput('function New-Panel { $a = New-Object System.Windows.Controls.Button; $a.SetResourceReference([System.Windows.Controls.Control]::StyleProperty, ''DialogButton''); $b = New-Object System.Windows.Controls.Button }', [ref]$null, [ref]$null)
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
        Get-CodeStyleKey -Function $fn -Variable 'a' | Should -Be 'DialogButton'
        Get-CodeStyleKey -Function $fn -Variable 'b' | Should -Be ''
    }

    It 'themes the sidebar buttons through SuiteCommon at start and at each toggle' {
        $script:ShellText | Should -Match '(?m)^\s+Set-ButtonTheme\s+-IsDark\s+\$true'
        $script:ShellText | Should -Match '(?m)^\s+Set-ButtonTheme\s+-IsDark\s+\$false'
    }

    It 'defines the dialog styles once, in MainWindow.xaml, and in no dialog' {
        $shell = $script:ShellText
        $shell | Should -Not -Match 'x:Key="(DialogButton|DialogAccentButton|ThemedButtonStyle|InlineButtonStyle|GridRowStyle|GridCellStyle)"'
        ([regex]::Matches($shell, [regex]::Escape($script:SharedStyleMarker))).Count | Should -BeGreaterOrEqual 4
        ([regex]::Matches($shell, 'Replace\(\$script:SharedStyleMarker')).Count | Should -BeGreaterOrEqual 4
        $shell | Should -Match 'function Get-SharedStyleXaml'
        $keys = [regex]::Match($shell, '\$keys = ([^\r\n]+)').Groups[1].Value
        foreach ($k in $script:SharedStyleKey) { $keys | Should -Match ([regex]::Escape("'$k'")) }
        ([regex]::Matches($keys, "'[A-Za-z]+'")).Count | Should -Be $script:SharedStyleKey.Count
    }

    It 'documents the one secondary style, and gives it a gray fill, readable theme text, a visible border and the size of the primary style' {
        $mainText = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'MainWindow.xaml') -Raw
        $mainText | Should -Match 'The one secondary style'
        [xml]$x = $mainText
        function Get-StyleSetter($key) {
            $st = $x.SelectSingleNode(("//*[local-name()='Style'][@*[local-name()='Key']='{0}']" -f $key))
            $st | Should -Not -BeNullOrEmpty -Because $key
            $h = @{}
            foreach ($se in @($st.SelectNodes("*[local-name()='Setter']"))) { $h[$se.GetAttribute('Property')] = Get-SetterValue $se }
            $h
        }
        $sec = Get-StyleSetter 'DialogButton'
        $pri = Get-StyleSetter 'DialogAccentButton'
        $sec['Background'] | Should -Match 'MahApps\.Brushes\.Gray'
        $sec['Foreground'] | Should -Match 'MahApps\.Brushes\.ThemeForeground'
        $sec['BorderBrush'] | Should -Match 'MahApps\.Brushes\.ThemeForeground'
        $sec['Height'] | Should -Be $pri['Height']
        $sec['MinWidth'] | Should -Be $pri['MinWidth']
        Test-WhiteBrush $pri['Foreground'] | Should -BeTrue
        # The inline button is the secondary style at another height. It sets no color of its own.
        $inl = Get-StyleSetter 'InlineButtonStyle'
        @($inl.Keys | Sort-Object) | Should -Be @('FontSize', 'Height')
    }

    It 'sets no color literal in any Style or ControlTemplate of <_.Source>' -ForEach $script:XamlDocsD {
        [xml]$x = $_.Text
        $nodes = @($x.SelectNodes('//*[local-name()="Style" or local-name()="ControlTemplate"]'))
        @(Get-HardColor -Nodes $nodes -Label $_.Source) | Should -BeNullOrEmpty
    }

    It 'scans the styles and templates of the main window (the scan is not empty)' {
        $doc = @($script:XamlDocs | Where-Object { $_.Source -eq 'MainWindow.xaml' })[0]
        [xml]$x = $doc.Text
        @($x.SelectNodes('//*[local-name()="Style"]')).Count | Should -BeGreaterThan 8
        @($x.SelectNodes('//*[local-name()="ControlTemplate"]')).Count | Should -BeGreaterThan 0
    }

    It 'changes no opacity in a hover, pressed, disabled, focus or selected state of <_.Source>' -ForEach $script:XamlDocsD {
        @(Get-OpacityInStateTrigger -Text $_.Text) | Should -BeNullOrEmpty
    }

    It 'detects an opacity in a state trigger (the check itself)' {
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
        $bad = "<ControlTemplate $ns TargetType=`"Button`"><ControlTemplate.Triggers><Trigger Property=`"IsMouseOver`" Value=`"True`"><Setter TargetName=`"Bd`" Property=`"Opacity`" Value=`"0.85`"/></Trigger></ControlTemplate.Triggers></ControlTemplate>"
        $ok = "<ControlTemplate $ns TargetType=`"Button`"><ControlTemplate.Triggers><Trigger Property=`"IsMouseOver`" Value=`"True`"><Setter TargetName=`"Bd`" Property=`"BorderBrush`" Value=`"{DynamicResource MahApps.Brushes.ThemeForeground}`"/></Trigger></ControlTemplate.Triggers></ControlTemplate>"
        @(Get-OpacityInStateTrigger -Text $bad).Count | Should -Be 1
        @(Get-OpacityInStateTrigger -Text $ok).Count | Should -Be 0
    }

    It 'shows hover as a border, pressed as swapped theme colors, and disabled as gray text in the button template and styles' {
        $doc = @($script:XamlDocs | Where-Object { $_.Source -eq 'MainWindow.xaml' })[0]
        [xml]$x = $doc.Text
        $tpl = $x.SelectSingleNode("//*[local-name()='ControlTemplate'][@*[local-name()='Key']='ThemedButtonTemplate']")
        $tpl | Should -Not -BeNullOrEmpty
        $hover = $tpl.SelectSingleNode(".//*[local-name()='Trigger'][@Property='IsMouseOver']/*[local-name()='Setter']")
        $hover.GetAttribute('Property') | Should -Be 'BorderBrush'
        $hover.GetAttribute('Value') | Should -Match 'ThemeForeground'
        $pressed = @($tpl.SelectNodes(".//*[local-name()='Trigger'][@Property='IsPressed']/*[local-name()='Setter']") | ForEach-Object { $_.GetAttribute('Value') }) -join ' '
        $pressed | Should -Match 'ThemeForeground'
        $pressed | Should -Match 'ThemeBackground'
        foreach ($key in 'ThemedButtonStyle', 'DialogAccentButton') {
            $st = $x.SelectSingleNode(("//*[local-name()='Style'][@*[local-name()='Key']='{0}']" -f $key))
            $dis = @($st.SelectNodes(".//*[local-name()='Trigger'][@Property='IsEnabled']/*[local-name()='Setter'][@Property='Foreground']"))
            $dis.Count | Should -Be 1 -Because $key
            $dis[0].GetAttribute('Value') | Should -Match 'Gray\.MouseOver' -Because $key
        }
        $doc.Text | Should -Not -Match 'change only the opacity'
    }

    It 'documents each known theme limit with a reason, and the limit holds' {
        . "$PSScriptRoot\KnownThemeLimits.ps1"
        @($script:KnownThemeLimit).Count | Should -BeGreaterOrEqual 1
        foreach ($l in $script:KnownThemeLimit) {
            foreach ($field in 'Id', 'Where', 'Reason', 'Source') { [string]$l[$field] | Should -Not -BeNullOrEmpty -Because ('{0} needs {1}' -f $l.Id, $field) }
            ([string]$l.Reason).Length | Should -BeGreaterOrEqual 60 -Because ('{0} needs a reason that says why' -f $l.Id)
        }
        $limit = @($script:KnownThemeLimit | Where-Object { $_.Id -eq 'sidebar-disabled-white-on-blue' })
        $limit.Count | Should -Be 1
        # The limit stays inert while the app never disables a sidebar button, and while SuiteCommon sets the local foreground.
        $shell = Get-Content -LiteralPath $script:ShellPath -Raw
        $shell | Should -Not -Match '\$(btnApps|btnPackages|btnTaskSequences|btnSUG|btnRings|btnOptions)\.IsEnabled'
        $suite = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1') -Raw
        $suite | Should -Match '\.Foreground\s*=\s*\$OnBlue'
    }

    It 'detects a color literal in a style, a template and an attribute (the check itself)' {
        [xml]$x = '<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><Style x:Key="S" TargetType="Button"><Setter Property="Foreground" Value="White"/><Setter Property="Background"><Setter.Value><SolidColorBrush Color="#1E1E1E"/></Setter.Value></Setter><Setter Property="BorderBrush" Value="{DynamicResource MahApps.Brushes.Accent}"/></Style><ControlTemplate x:Key="T" TargetType="Button"><Border BorderBrush="#555555" Background="{TemplateBinding Background}"><Border.Background><SolidColorBrush Color="#123456"/></Border.Background></Border></ControlTemplate></ResourceDictionary>'
        $nodes = @($x.SelectNodes('//*[local-name()="Style" or local-name()="ControlTemplate"]'))
        $found = @(Get-HardColor -Nodes $nodes -Label 'sample')
        $found.Count | Should -Be 4
        ($found -join ' ') | Should -Match '#123456'
        $ok = '<ControlTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" TargetType="Button"><Border><Border.Background><SolidColorBrush Color="{DynamicResource MahApps.Colors.AccentBase}"/></Border.Background></Border></ControlTemplate>'
        [xml]$y = $ok
        @(Get-HardColor -Nodes @($y.DocumentElement) -Label 'ok').Count | Should -Be 0
        ($found -join ' ') | Should -Match 'White'
        ($found -join ' ') | Should -Match '#1E1E1E'
        ($found -join ' ') | Should -Match '#555555'
    }
}

Describe 'Every test file requires Pester 5.1.0' {
    It 'has the #Requires line in <_.Name>' -ForEach @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' -File -Recurse | ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } }) {
        (Get-Content -LiteralPath $Path -TotalCount 5) -join "`n" | Should -Match "(?m)^#Requires -Modules @\{ ModuleName = 'Pester'; ModuleVersion = '5\.1\.0' \}"
    }

    It 'finds the test files (the scan is not empty)' {
        @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' -File -Recurse).Count | Should -BeGreaterOrEqual 6
    }
}

Describe 'No raw exception text in the UI (f)' {
    It 'flags raw exception text that leaves the catch block (the check itself)' {
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Add-LogLine ("x {0}" -f $_.Exception.Message) }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $row.Detail = $_.Exception.Message }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $State.ErrorMsg = $_.Exception.Message }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { return @{ Success = $false; Error = $_.ToString() } }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { return [pscustomobject]@{ ErrorMessage = $_.Exception.Message } }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $errors += $_.Exception.Message }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $msg = $_.Exception.Message; Add-LogLine $msg }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { throw $_ }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Show-ThemedMessage -Message ("x: {0}" -f $err.Exception.Message) }').Count | Should -Be 1
    }

    It 'allows the log file, the helper, a condition and a deliberate swallow (the check itself)' {
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Write-Log ("x {0}" -f $_.Exception.Message) -Level ERROR; Add-LogLine "Failed. See the log file." }').Count | Should -Be 0
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Write-Log "x: $_" -Level WARN }').Count | Should -Be 0
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { return @{ Error = (Get-UserErrorText -ErrorRecord $_ -Fallback "Failed.") } }').Count | Should -Be 0
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { if ($_.Exception.Message -match "already") { 1 } else { 2 } }').Count | Should -Be 0
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $null = $_ }').Count | Should -Be 0
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $x = @(1, 2) | ForEach-Object { $_ + 1 } }').Count | Should -Be 0
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { throw }').Count | Should -Be 0
    }

    It 'scans a nested catch block on its own (the check itself)' {
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { try { 2 } catch { Add-LogLine $_ } }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { try { 2 } catch { $null = $_ }; Write-Log $_ }').Count | Should -Be 0
    }

    It 'finds the catch blocks of <Name> (the scan is not empty)' -ForEach $script:AppPsFilesD {
        (Get-CatchCount -Path $Path) | Should -BeGreaterOrEqual 20
    }

    It 'sends no raw exception text to the UI from <Name>' -ForEach $script:AppPsFilesD {
        @(Get-RawExceptionToUi -Path $Path) | Should -BeNullOrEmpty
    }

    It 'keeps the raw detail in ErrorDetail, and only the audit record reads it' {
        $shellLines = @(Get-Content -LiteralPath $script:ShellPath | Where-Object { $_ -match 'ErrorDetail' })
        $shellLines.Count | Should -BeGreaterOrEqual 4
        foreach ($l in $shellLines) { $l | Should -Match '^\s*Result\s*=' }
        $moduleLines = @(Get-Content -LiteralPath $script:ModulePath | Where-Object { $_ -match 'ErrorDetail' -and $_ -notmatch 'function Get-ErrorDetailText|ErrorDetail \(|a result carries it in ErrorDetail' })
        $moduleLines.Count | Should -BeGreaterOrEqual 5
        foreach ($l in $moduleLines) { $l | Should -Match 'ErrorDetail\s*=\s*\(Get-ErrorDetailText -ErrorRecord \$_\)|^\s*Result\s*=' }
    }

    It 'uses Get-UserErrorText in the start script and in the module, and exports it' {
        (Get-Content -LiteralPath $script:ShellPath -Raw) | Should -Match 'Get-UserErrorText'
        (Get-Content -LiteralPath $script:ModulePath -Raw) | Should -Match 'function Get-UserErrorText'
        (Get-Content -LiteralPath (Join-Path $script:RepoRoot 'Module\DeploymentHelperCommon.psd1') -Raw) | Should -Match "'Get-UserErrorText'"
    }
}
