#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Tooltip guard. It parses the XAML file and the XAML here-strings of the start script. It creates no WPF control,
# opens no window and needs no ConfigMgr connection.
#
# Rule: every control in the audited list has a ToolTip attribute that is not empty, is ASCII only, ends in a period
# and is not longer than 100 characters, in one sentence (the same limit as Theme.Guard.Tests.ps1). Add the x:Name of a new control to the list when it gets a tooltip.

BeforeDiscovery {
    # The x:Name of every audited control, grouped by the file that holds it.
    $script:TtAudited = @{
        'MainWindow.xaml'            = @(
            'toggleTheme', 'btnApps', 'btnPackages', 'btnTaskSequences', 'btnSUG', 'btnRings', 'btnOptions',
            'cboApplyTemplate', 'txtTargetName', 'btnBrowseTarget', 'cboProgram', 'txtCollection', 'btnBrowseCollection',
            'radAvailable', 'radRequired', 'dtpAvailable', 'dtpDeadline', 'cboNotification', 'radLocalTime', 'radUtc',
            'btnDPPicker', 'btnDistributeContent',
            'chkOverrideSW', 'chkRebootOutSW', 'chkMetered',
            'chkPkgOverrideSW', 'chkPkgRebootOutSW', 'chkPkgMetered', 'cboPkgFastNetwork', 'cboPkgSlowNetwork', 'cboPkgRerun',
            'chkTsOverrideSW', 'chkTsRebootOutSW', 'chkTsMetered', 'chkShowProgress',
            'chkSugOverrideSW', 'chkSugAllowRestart', 'chkSugMetered', 'chkMSFallback', 'chkBoundaryFallback', 'chkFullScan',
            'btnValidate', 'btnDeploy',
            'cboRingType', 'txtRingObject', 'btnRingBrowseObject', 'cboRingProgram', 'cboRingPlan', 'btnRingReloadPlans', 'btnRingOpenFolder',
            'dtpRingStart', 'btnRingExpand', 'chkRingHold', 'btnRingPickCollection', 'btnRingValidate', 'btnRingCreate',
            'btnRingRefreshRuns', 'chkRingShowClosed',
            'btnRollout', 'btnRoAdd', 'btnRoImport', 'btnRoLookup', 'btnRoRemove', 'txtRoTemplate', 'txtRoFolder',
            'txtRoLimitServer', 'txtRoLimitWorkstation', 'txtRoLimitTest', 'txtRoTestCollection', 'txtRoRefresh', 'txtRoDelay',
            'txtRoTestDeadline', 'txtRoProdDeadline', 'cboRoNotification', 'chkRoDryRun', 'cboRoRuns', 'btnRoOpenRun',
            'btnRoPreview', 'btnRoStep1', 'btnRoStep2'
        )
        'start-deploymenthelper.ps1' = @('txtFilter', 'btnRefresh', 'btnSelectAll', 'btnSelectNone', 'txtList', 'btnFile', 'btnCheck')
    }
    $script:TtCases = @()
    foreach ($file in $script:TtAudited.Keys) {
        foreach ($name in $script:TtAudited[$file]) { $script:TtCases += @{ File = $file; Name = $name } }
    }
    $script:TtSyncCases = @(@{ Listed = @($script:TtAudited['MainWindow.xaml']) })
}

BeforeAll {
    $script:TtRepoRoot = Split-Path -Parent $PSScriptRoot
    $script:TtNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    $script:TtKeyNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml'
    $script:TtMaxLength = 100

    # The XAML documents of one file: the file itself, or the XAML here-strings of a script.
    function Get-TtDocument {
        param([string]$RelativePath)
        $text = Get-Content -LiteralPath (Join-Path $script:TtRepoRoot $RelativePath) -Raw
        if ($RelativePath -like '*.xaml') { return @($text) }
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$errors)
        $nodes = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.StringConstantType -in 'SingleQuotedHereString', 'DoubleQuotedHereString' }, $true)
        $docs = @()
        foreach ($n in @($nodes)) {
            if ($n.Value -match ('^\s*<[^>]*xmlns="' + [regex]::Escape($script:TtNamespace) + '"')) { $docs += $n.Value }
        }
        return $docs
    }

    # The elements with the x:Name over all documents of the file, with their ToolTip attribute (empty when there is none).
    function Get-TtElement {
        param([string]$RelativePath, [string]$Name)
        $found = @()
        foreach ($doc in @(Get-TtDocument -RelativePath $RelativePath)) {
            [xml]$x = $doc
            foreach ($el in @($x.SelectNodes('//*'))) {
                if ($el.GetAttribute('Name', $script:TtKeyNamespace) -eq $Name) {
                    $found += , [PSCustomObject]@{ Element = $el.LocalName; ToolTip = $el.GetAttribute('ToolTip') }
                }
            }
        }
        return $found
    }
}

Describe 'Tooltips of the audited controls' {
    It 'gives <Name> in <File> a tooltip' -ForEach $script:TtCases {
        $hits = @(Get-TtElement -RelativePath $File -Name $Name)
        $hits.Count | Should -BeGreaterThan 0 -Because ('{0} must exist in {1}' -f $Name, $File)
        foreach ($h in $hits) {
            $h.ToolTip | Should -Not -BeNullOrEmpty -Because ('{0} needs a ToolTip' -f $Name)
            $h.ToolTip | Should -Match '^[\x20-\x7E]+$' -Because ('the tooltip of {0} must be ASCII only' -f $Name)
            $h.ToolTip | Should -Match '\.$' -Because ('the tooltip of {0} must end in a period' -f $Name)
            $h.ToolTip.TrimEnd('.') | Should -Not -Match '\.\s' -Because ('the tooltip of {0} must be one sentence' -f $Name)
            $h.ToolTip.Length | Should -BeLessOrEqual $script:TtMaxLength -Because ('the tooltip of {0} is too long' -f $Name)
        }
    }

    It 'lists every control of MainWindow.xaml that has a literal tooltip' -ForEach $script:TtSyncCases {
        [xml]$x = Get-Content -LiteralPath (Join-Path $script:TtRepoRoot 'MainWindow.xaml') -Raw
        $missing = @()
        foreach ($el in @($x.SelectNodes('//*[@ToolTip]'))) {
            $n = $el.GetAttribute('Name', $script:TtKeyNamespace)
            if ($n -and $n -notin $Listed -and $el.GetAttribute('ToolTip') -notmatch '^\{') { $missing += $n }
        }
        $missing | Should -BeNullOrEmpty -Because 'add every named control that has a literal tooltip to the audited list'
    }
}
