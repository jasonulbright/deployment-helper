# Helpers of Theme.Guard.Tests.ps1. Dot-source from BeforeDiscovery and from BeforeAll.
# They parse files. They create no WPF control and open no window.

$script:RepoRoot = Split-Path -Parent $PSScriptRoot
$script:ThemeProperty = @('FontSize', 'FontWeight', 'FontStyle', 'TextDecorations', 'FontFamily', 'FontStretch')
$script:BannedPattern = '(?i)\b(easily|simply|seamless(ly)?|powerful|robust|streamline[sd]?|streamlining|quickly|just|effortless(ly)?|intuitive(ly)?|handy|helps you|allows you to)\b'
$script:XamlNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'
$script:XamlKeyNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml'
$script:ExcludedFolder = '^(\.claude|\.git)([\\/]|$)'

# The path of a file relative to the repository root. A worktree lives under .claude, so a filter
# on the full path would match every file. A filter on the relative path does not.
function Get-RelativeRepoPath {
    param([string]$FullName)
    return $FullName.Substring($script:RepoRoot.Length).TrimStart('\', '/')
}

# True when a brush value names an accent or a blue color.
function Test-BlueBrush {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    if ($Value -match 'MahApps\.(Brushes|Colors)\.(Accent\d*|AccentBase|Highlight|Button\.Flat\.Background\.Pressed)\b') { return $true }
    if ($Value -match '(?i)^\s*(blue|dodgerblue|royalblue|steelblue|navy|mediumblue|darkblue|cornflowerblue)\s*$') { return $true }
    if ($Value -match '^\s*#(?:[0-9A-Fa-f]{2})?([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})\s*$') {
        $r = [Convert]::ToInt32($Matches[1], 16) / 255.0
        $g = [Convert]::ToInt32($Matches[2], 16) / 255.0
        $b = [Convert]::ToInt32($Matches[3], 16) / 255.0
        $max = [Math]::Max($r, [Math]::Max($g, $b)); $min = [Math]::Min($r, [Math]::Min($g, $b))
        $d = $max - $min
        if ($d -lt 0.2 -or $max -eq 0) { return $false }
        if ($max -ne $b) { return $false }
        $hue = 60.0 * ((4.0 + (($r - $g) / $d)))
        return ($hue -ge 190 -and $hue -le 260)
    }
    return $false
}

function Test-WhiteBrush {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value -match '(?i)^\s*(white|#FFFFFF|#FFFFFFFF)\s*$' -or $Value -match 'MahApps\.Brushes\.(Ideal|ThemeIdeal)Foreground')
}

# The value of a Setter: the Value attribute, or the Color of a nested brush.
function Get-SetterValue {
    param($Setter)
    if ($Setter.HasAttribute('Value')) { return $Setter.GetAttribute('Value') }
    $node = $Setter.SelectSingleNode('.//*[@Color]')
    if ($node) { return $node.GetAttribute('Color') }
    return ''
}

# The script files of the app: the start script and the module. The tests, the tools and the vendored library are outside.
function Get-AppScriptFile {
    $files = @(Get-ChildItem -LiteralPath $script:RepoRoot -File -Recurse | Where-Object {
            $rel = Get-RelativeRepoPath $_.FullName
            $_.Extension -in '.ps1', '.psm1' -and $_.Name -notlike '*.Tests.ps1' -and $rel -notmatch '^(\.claude|\.git|Tests|tools|Lib|Logs|Rings|Templates)([\\/]|$)'
        })
    return $files
}

# The styles that the dialogs take from MainWindow.xaml. The start script holds the same list (Get-SharedStyleXaml).
$script:SharedStyleMarker = '<!--SHARED_STYLES-->'
$script:SharedStyleKey = @('ThemedButtonTemplate', 'ThemedButtonStyle', 'DialogAccentButton', 'DialogButton', 'InlineButtonStyle', 'GridRowStyle', 'GridCellStyle')
function Get-SharedStyleXaml {
    [xml]$main = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'MainWindow.xaml') -Raw
    $nodes = @($main.SelectNodes("//*[local-name()='ResourceDictionary']/*[local-name()='ControlTemplate' or local-name()='Style']") | Where-Object {
            $key = $_.GetAttribute('Key', $script:XamlKeyNamespace)
            ($key -in $script:SharedStyleKey) -or (-not $key -and $_.GetAttribute('TargetType') -in 'ComboBoxItem', 'ListBoxItem')
        })
    return (($nodes | ForEach-Object { $_.OuterXml }) -join [Environment]::NewLine)
}

# Every XAML document of the app: the XAML files and the XAML here-strings of the scripts.
# A here-string counts as XAML when its first tag declares the WPF namespace.
function Get-XamlDocument {
    $docs = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $script:RepoRoot -Filter '*.xaml' -File -Recurse | Where-Object { (Get-RelativeRepoPath $_.FullName) -notmatch $script:ExcludedFolder })) {
        $docs += [PSCustomObject]@{ Source = $f.Name; Text = (Get-Content -LiteralPath $f.FullName -Raw) }
    }
    foreach ($f in @(Get-AppScriptFile)) {
        $raw = Get-Content -LiteralPath $f.FullName -Raw
        $i = 0
        foreach ($m in [regex]::Matches($raw, "(?s)@['""]\r?\n(\s*<[^\r\n]*\r?\n(?:[^\r\n]*\r?\n)*?)['""]@")) {
            $text = $m.Groups[1].Value
            if ($text -notmatch ('^\s*<[^>]*xmlns="' + [regex]::Escape($script:XamlNamespace) + '"')) { continue }
            $i++
            # The dialogs take the shared styles from MainWindow.xaml, as the start script does at run time.
            $text = $text.Replace($script:SharedStyleMarker, (Get-SharedStyleXaml))
            $docs += [PSCustomObject]@{ Source = ('{0} here-string {1}' -f $f.Name, $i); Text = $text }
        }
    }
    return $docs
}

# Returns one text per blue-background rule that has no white foreground.
function Get-BlueWithoutWhite {
    param([string]$Text, [string]$Source)
    [xml]$x = $Text
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($el in @($x.SelectNodes('//*'))) {
        if ($el.LocalName -eq 'Setter' -and $el.GetAttribute('Property') -match '(^|\.)Background$') {
            if (Test-BlueBrush (Get-SetterValue $el)) {
                $fg = $el.ParentNode.ChildNodes | Where-Object { $_.LocalName -eq 'Setter' -and $_.GetAttribute('Property') -eq 'Foreground' } | Select-Object -First 1
                if (-not $fg -or -not (Test-WhiteBrush (Get-SetterValue $fg))) {
                    [void]$found.Add(('{0}: Setter Background {1} in <{2}> has no white Foreground' -f $Source, (Get-SetterValue $el), $el.ParentNode.LocalName))
                }
            }
        }
        elseif ($el.HasAttribute('Background') -and $el.LocalName -ne 'Setter') {
            if (Test-BlueBrush $el.GetAttribute('Background')) {
                if (-not (Test-WhiteBrush $el.GetAttribute('Foreground'))) {
                    [void]$found.Add(('{0}: <{1}> Background {2} has no white Foreground' -f $Source, $el.LocalName, $el.GetAttribute('Background')))
                }
            }
        }
    }
    return $found.ToArray()
}

# Every string of a script: the constant strings and the text parts of expandable strings.
function Get-ScriptString {
    param([string]$Path)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $nodes = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true)
    foreach ($n in $nodes) {
        # An attribute such as SuppressMessage holds a justification for the analyzer, not a UI string.
        $p = $n.Parent; $inAttribute = $false
        while ($p) { if ($p -is [System.Management.Automation.Language.AttributeAst]) { $inAttribute = $true; break }; $p = $p.Parent }
        if (-not $inAttribute) { [string]$n.Value }
    }
}

# The text a user can read in a XAML document: tooltips and text attributes.
function Get-XamlUiText {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($el in @($x.SelectNodes('//*'))) {
        foreach ($a in @($el.Attributes)) {
            if ($a.LocalName -match '^(ToolTip|Text|Content|Header|Title|Watermark|OnContent|OffContent|ToolTipService.*)$' -or $a.Name -match 'Watermark$') {
                if ($a.Value -notmatch '^\{') { $a.Value }
            }
        }
    }
}

# The tooltips of a XAML document that are not a binding.
function Get-XamlToolTip {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($el in @($x.SelectNodes('//*[@ToolTip]'))) {
        $v = $el.GetAttribute('ToolTip')
        if ($v -notmatch '^\{') { $v }
    }
}

# True when a tooltip is one short sentence: it starts with a capital letter, has one full stop at the end, and no more
# than 100 characters. A full stop inside round brackets or after "e.g" does not count as a sentence end.
function Test-OneSentence {
    param([string]$Text)
    if ($Text.Length -gt 100) { return $false }
    if ($Text -cnotmatch '^[A-Z]') { return $false }
    if ($Text -notmatch '\.$') { return $false }
    return ($Text.TrimEnd('.') -notmatch '\.\s')
}

# True when every word of a button text starts with a capital letter or a digit. Short words in the middle may stay small.
function Test-TitleCase {
    param([string]$Text)
    $small = 'a', 'an', 'and', 'as', 'at', 'by', 'for', 'from', 'in', 'of', 'on', 'or', 'the', 'to', 'with'
    $words = @(($Text -replace '\.\.\.$', '').Trim() -split '\s+' | Where-Object { $_ })
    for ($i = 0; $i -lt $words.Count; $i++) {
        if ($i -gt 0 -and $small -contains $words[$i]) { continue }
        if ($words[$i] -cnotmatch '^[A-Z0-9]') { return $false }
    }
    return $words.Count -gt 0
}

# One text per catch block statement that sends exception text anywhere but to the log file.
# The exception variable ($_, $PSItem, $Error, $ex, $err) may appear in a catch block only as an argument of:
#   - Write-Log or Write-LogErrorRecord (the log file),
#   - Get-UserErrorText (it returns the message of an own throw, and a fixed sentence for any other error),
#   - the condition of an if statement (a decision, not an output),
#   - $null = $_ (a deliberate swallow).
# A pipeline script block inside the catch block rebinds $_, so it is not scanned. A nested catch block is scanned on its own.
# This covers sinks (Add-LogLine, dialogs), properties (Text, Error, ErrorMsg), hashtable keys, variables and throw $_.
function Get-RawExceptionToUi {
    param([string]$Path, [string]$ScriptText)
    $ast = if ($ScriptText) { [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$null, [ref]$null) } else { [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null) }
    $errorName = @('_', 'PSItem', 'Error', 'ex', 'err')
    $safeCommand = @('Write-Log', 'Write-LogErrorRecord', 'Get-UserErrorText')
    $found = @()
    foreach ($c in @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CatchClauseAst] }, $true))) {
        $vars = @($c.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] -and $errorName -contains $n.VariablePath.UserPath }, $true))
        foreach ($v in $vars) {
            $allowed = $false; $skip = $false
            $q = $v
            $statement = $null
            while ($q -and $q -ne $c) {
                if ($q -is [System.Management.Automation.Language.CatchClauseAst]) { $skip = $true; break }
                if ($q -is [System.Management.Automation.Language.ScriptBlockExpressionAst] -and @('_', 'PSItem') -contains $v.VariablePath.UserPath) { $skip = $true; break }
                if ($q -is [System.Management.Automation.Language.CommandAst] -and $safeCommand -contains [string]$q.GetCommandName()) { $allowed = $true; break }
                if ($q -is [System.Management.Automation.Language.AssignmentStatementAst] -and $q.Left.Extent.Text -eq '$null') { $allowed = $true; break }
                if ($q.Parent -is [System.Management.Automation.Language.IfStatementAst]) {
                    $parentIf = $q.Parent
                    $isCondition = $false
                    foreach ($cl in $parentIf.Clauses) { if ($cl.Item1 -eq $q) { $isCondition = $true } }
                    if ($isCondition) { $allowed = $true; break }
                }
                if ($q -is [System.Management.Automation.Language.StatementAst] -and -not $statement) { $statement = $q }
                $q = $q.Parent
            }
            if ($skip -or $allowed) { continue }
            $where = if ($statement) { $statement } else { $v }
            $text = $where.Extent.Text
            if ($text.Length -gt 120) { $text = $text.Substring(0, 120) }
            $found += ('line {0}: {1}' -f $where.Extent.StartLineNumber, $text)
        }
    }
    return @($found | Select-Object -Unique)
}

# The number of catch blocks in a script.
function Get-CatchCount {
    param([string]$Path)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    return @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CatchClauseAst] }, $true)).Count
}

# A function of a script, found by name.
function Get-ScriptFunction {
    param([string]$Path, [string]$Name)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    return $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
}

# The assignments of the form $name.Member = value in an AST node, as a hashtable of member name to value text.
function Get-MemberAssignment {
    param($Node, [string]$Variable)
    $h = @{}
    foreach ($a in @($Node.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and $n.Left.Expression -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.Left.Expression.VariablePath.UserPath -eq $Variable }, $true))) {
        $h[[string]$a.Left.Member.Value] = $a.Right.Extent.Text
    }
    return $h
}

# The controls of one type that a script creates with New-Object, with their variable names and the function that holds them.
function Get-CodeControl {
    param([string]$Path, [string]$TypeName)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $pattern = 'New-Object\s+(-TypeName\s+)?System\.Windows\.Controls\.' + [regex]::Escape($TypeName) + '\b'
    foreach ($a in @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.Right.Extent.Text -match $pattern }, $true))) {
        $fn = $a.Parent
        while ($fn -and $fn -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $fn = $fn.Parent }
        if (-not $fn) { $fn = $ast }
        [PSCustomObject]@{ Variable = $a.Left.VariablePath.UserPath; Line = $a.Extent.StartLineNumber; Function = $fn; FunctionName = $(if ($fn -is [System.Management.Automation.Language.FunctionDefinitionAst]) { $fn.Name } else { '' }) }
    }
}

# The style key that a code-built control takes with SetResourceReference([Control]::StyleProperty, 'Key'), or ''.
function Get-CodeStyleKey {
    param($Function, [string]$Variable)
    $calls = @($Function.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and [string]$n.Member.Value -eq 'SetResourceReference' -and $n.Expression -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.Expression.VariablePath.UserPath -eq $Variable -and $n.Arguments.Count -eq 2 -and $n.Arguments[0].Extent.Text -match 'StyleProperty' }, $true))
    if ($calls.Count -eq 0) { return '' }
    return [string]$calls[0].Arguments[1].Value
}

# Lists every hard-coded color in the Setters, the attributes and the nested brushes of the given nodes.
# Only a DynamicResource or a TemplateBinding is allowed. A property is a brush property when its name ends in
# Brush, Background, Foreground, Fill or Stroke, also as Border.Background or as an attached name.
# Any Color attribute that is not a DynamicResource or a TemplateBinding is a literal, wherever it sits.
function Get-HardColor {
    param($Nodes, [string]$Label)
    $brushProp = '(^|[.:])\w*(Brush|Background|Foreground|Fill|Stroke)$'
    $allowed = '^\{(DynamicResource|TemplateBinding)\b'
    foreach ($n in $Nodes) {
        foreach ($setter in @($n.SelectNodes('descendant-or-self::*[local-name()="Setter"]'))) {
            if ($setter.GetAttribute('Property') -notmatch $brushProp) { continue }
            $v = $setter.GetAttribute('Value')
            if ($v -and $v -notmatch $allowed) { '{0}: Setter {1}="{2}"' -f $Label, $setter.GetAttribute('Property'), $v }
        }
        foreach ($el in @($n.SelectNodes('descendant-or-self::*[not(local-name()="Setter")]'))) {
            foreach ($attr in @($el.Attributes)) {
                if ($attr.Name -match '^xmlns' -or $attr.NamespaceURI -eq 'http://www.w3.org/2000/xmlns/') { continue }
                if ($attr.LocalName -match $brushProp -and $attr.Value -notmatch $allowed) { '{0}: {1} {2}="{3}"' -f $Label, $el.LocalName, $attr.LocalName, $attr.Value }
                elseif ($attr.LocalName -eq 'Color' -and $attr.Value -notmatch $allowed) { '{0}: {1} holds the color {2}' -f $Label, $el.LocalName, $attr.Value }
            }
        }
    }
}

# One text per trigger that sets a font property and whose condition names a theme (Dark, Light, Theme).
# A trigger on a control state, such as the selected tab, sets the same font in both themes and is allowed.
function Get-ThemeFontTrigger {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($trigger in @($x.SelectNodes("//*[local-name()='Trigger' or local-name()='DataTrigger' or local-name()='MultiTrigger' or local-name()='MultiDataTrigger']"))) {
        $sets = @($trigger.SelectNodes("*[local-name()='Setter']") | Where-Object { $script:ThemeProperty -contains $_.GetAttribute('Property') })
        if ($sets.Count -eq 0) { continue }
        $condition = (@($trigger.Attributes | ForEach-Object { $_.Value }) + @($trigger.SelectNodes(".//*[local-name()='Condition']") | ForEach-Object { @($_.Attributes | ForEach-Object { $_.Value }) })) -join ' '
        if ($condition -match '(?i)IsDark|DarkMode|IsLight|LightMode|ThemeMode|\bTheme\b|\bDark\b') { $trigger.OuterXml.Substring(0, [Math]::Min(100, $trigger.OuterXml.Length)) }
    }
}

# The keys of a style and of the styles that it is based on, in order, from a XAML document.
function Get-StyleChain {
    param($Doc, [string]$Key)
    $chain = @()
    $guard = 0
    while ($Key -and $guard -lt 10) {
        $guard++
        $chain += $Key
        $st = $Doc.SelectSingleNode(("//*[local-name()='Style'][@*[local-name()='Key']='{0}']" -f $Key))
        if (-not $st) { break }
        $based = $st.GetAttribute('BasedOn')
        $Key = if ($based -match '(?:Static|Dynamic)Resource\s+([^\}\s]+)') { $Matches[1] } else { '' }
    }
    return $chain
}

# The style chain of a Button: the Style attribute, or the style that sits inside the Button.
function Get-ButtonStyleChain {
    param($Doc, $Button)
    $style = $Button.GetAttribute('Style')
    if ($style -match '(?:Static|Dynamic)Resource\s+([^\}\s]+)') { return @(Get-StyleChain -Doc $Doc -Key $Matches[1]) }
    $inline = $Button.SelectSingleNode("*[local-name()='Button.Style']/*[local-name()='Style']")
    if ($inline) {
        $based = $inline.GetAttribute('BasedOn')
        if ($based -match '(?:Static|Dynamic)Resource\s+([^\}\s]+)') { return @('(inline)') + @(Get-StyleChain -Doc $Doc -Key $Matches[1]) }
        return @('(inline)')
    }
    return @()
}

# One text per Button that has no Style, or a Style that is not based on a MahApps button style.
function Get-ButtonWithoutSharedStyle {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($b in @($x.SelectNodes("//*[local-name()='Button']"))) {
        $chain = @(Get-ButtonStyleChain -Doc $x -Button $b)
        $based = $false
        foreach ($key in $chain) {
            $st = $x.SelectSingleNode(("//*[local-name()='Style'][@*[local-name()='Key']='{0}']" -f $key))
            if ($st -and $st.GetAttribute('BasedOn') -match 'MahApps\.Styles\.Button') { $based = $true }
        }
        if (-not $based) { 'Button {0} has no shared style' -f $b.GetAttribute('Name', $script:XamlKeyNamespace) }
    }
}
