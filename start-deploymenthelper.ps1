<#
.SYNOPSIS
    WinForms front-end for Deployment Helper - safe MECM application deployment.

.DESCRIPTION
    Single-pane deployment tool with pre-execution validation, safety guardrails,
    and immutable deployment audit logging.

    Features:
      - Enter application or SUG, collection
      - 5-check validation engine (app exists, content distributed, collection valid,
        collection safe, no duplicate deployment)
      - Deployment templates for consistent configurations
      - Immutable JSONL deployment log
      - CSV and HTML history export
      - Dark mode / light mode

.EXAMPLE
    .\start-deploymenthelper.ps1

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8+
      - Windows Forms (System.Windows.Forms)
      - Configuration Manager console installed

    ScriptName : start-deploymenthelper.ps1
    Purpose    : WinForms front-end for safe MECM application deployment
    Version    : 1.0.0
    Updated    : 2026-02-27
#>

param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

$moduleRoot = Join-Path $PSScriptRoot "Module"
Import-Module (Join-Path $moduleRoot "DeploymentHelperCommon.psd1") -Force

# Initialize tool logging
$toolLogFolder = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $toolLogFolder)) {
    New-Item -ItemType Directory -Path $toolLogFolder -Force | Out-Null
}
$toolLogPath = Join-Path $toolLogFolder ("DeploymentHelper-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $toolLogPath

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Set-ModernButtonStyle {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory)][System.Drawing.Color]$BackColor
    )

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $BackColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand

    $hover = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 18),
        [Math]::Max(0, $BackColor.G - 18),
        [Math]::Max(0, $BackColor.B - 18)
    )
    $down = [System.Drawing.Color]::FromArgb(
        [Math]::Max(0, $BackColor.R - 36),
        [Math]::Max(0, $BackColor.G - 36),
        [Math]::Max(0, $BackColor.B - 36)
    )

    $Button.FlatAppearance.MouseOverBackColor = $hover
    $Button.FlatAppearance.MouseDownBackColor = $down
}

function Enable-DoubleBuffer {
    param([Parameter(Mandatory)][System.Windows.Forms.Control]$Control)
    $prop = $Control.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
    if ($prop) { $prop.SetValue($Control, $true, $null) | Out-Null }
}

function Show-SearchDialog {
    param(
        [string]$Title = "Search",
        [string]$SearchLabel = "Search:",
        [string[]]$ColumnNames,
        [string]$NameColumn,
        [scriptblock]$SearchAction
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(600, 450)
    $dlg.MinimumSize = $dlg.Size
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.BackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(30, 30, 30) } else { [System.Drawing.Color]::FromArgb(245, 246, 248) }
    $dlg.ForeColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(220, 220, 220) } else { [System.Drawing.Color]::Black }

    $pnlTop = New-Object System.Windows.Forms.Panel
    $pnlTop.Dock = [System.Windows.Forms.DockStyle]::Top
    $pnlTop.Height = 40
    $pnlTop.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 4)
    $dlg.Controls.Add($pnlTop)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtSearch.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $txtSearch.BackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(45, 45, 45) } else { [System.Drawing.Color]::FromArgb(250, 250, 250) }
    $txtSearch.ForeColor = $dlg.ForeColor
    $txtSearch.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
    $pnlTop.Controls.Add($txtSearch)

    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Text = "Search"
    $btnSearch.Dock = [System.Windows.Forms.DockStyle]::Right
    $btnSearch.Width = 80
    $btnSearch.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSearch.FlatAppearance.BorderSize = 0
    $btnSearch.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnSearch.ForeColor = [System.Drawing.Color]::White
    $btnSearch.Cursor = [System.Windows.Forms.Cursors]::Hand
    $pnlTop.Controls.Add($btnSearch)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Dock = [System.Windows.Forms.DockStyle]::Left
    $lblHint.Text = $SearchLabel
    $lblHint.Width = 60
    $lblHint.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $pnlTop.Controls.Add($lblHint)

    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Dock = [System.Windows.Forms.DockStyle]::Fill
    $dgv.ReadOnly = $true
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $dgv.MultiSelect = $false
    $dgv.RowHeadersVisible = $false
    $dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $dgv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dgv.BackgroundColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(40, 40, 40) } else { [System.Drawing.Color]::White }
    $dgv.DefaultCellStyle.BackColor = $dgv.BackgroundColor
    $dgv.DefaultCellStyle.ForeColor = $dlg.ForeColor
    $dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $dgv.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $dgv.AlternatingRowsDefaultCellStyle.BackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(48, 48, 48) } else { [System.Drawing.Color]::FromArgb(248, 250, 252) }
    $dgv.ColumnHeadersDefaultCellStyle.BackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(50, 50, 50) } else { [System.Drawing.Color]::FromArgb(240, 240, 240) }
    $dgv.ColumnHeadersDefaultCellStyle.ForeColor = $dlg.ForeColor
    $dgv.EnableHeadersVisualStyles = $false
    $dgv.GridColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(60, 60, 60) } else { [System.Drawing.Color]::FromArgb(230, 230, 230) }
    $dgv.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $dlg.Controls.Add($dgv)

    $pnlBottom = New-Object System.Windows.Forms.Panel
    $pnlBottom.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $pnlBottom.Height = 44
    $dlg.Controls.Add($pnlBottom)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Size = New-Object System.Drawing.Size(90, 30)
    $btnOK.Location = New-Object System.Drawing.Point(([int]($dlg.ClientSize.Width / 2 - 100)), 7)
    $btnOK.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
    $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOK.FlatAppearance.BorderSize = 0
    $btnOK.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $btnOK.ForeColor = [System.Drawing.Color]::White
    $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $pnlBottom.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(([int]($dlg.ClientSize.Width / 2 + 10)), 7)
    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(70, 70, 70) } else { [System.Drawing.Color]::FromArgb(200, 200, 200) }
    $btnCancel.BackColor = $dlg.BackColor
    $btnCancel.ForeColor = $dlg.ForeColor
    $btnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $pnlBottom.Controls.Add($btnCancel)

    $dlg.AcceptButton = $btnOK
    $dlg.CancelButton = $btnCancel

    # Column setup
    foreach ($col in $ColumnNames) {
        [void]$dgv.Columns.Add($col, $col)
    }

    # Search handler
    $doSearch = {
        $term = $txtSearch.Text.Trim()
        if ($term.Length -lt 2) { return }
        $dlg.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $dgv.Rows.Clear()
        $results = & $SearchAction $term
        foreach ($item in $results) {
            $row = @()
            foreach ($col in $ColumnNames) { $row += [string]$item.$col }
            [void]$dgv.Rows.Add($row)
        }
        $dlg.Cursor = [System.Windows.Forms.Cursors]::Default
    }

    $btnSearch.Add_Click($doSearch)
    $txtSearch.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.SuppressKeyPress = $true
            & $doSearch
        }
    })

    # Double-click row = OK
    $dgv.Add_CellDoubleClick({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK; $dlg.Close() })

    # Z-order
    $dgv.BringToFront()

    # Focus search box
    $dlg.Add_Shown({ $txtSearch.Focus() })

    $result = $dlg.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $dgv.SelectedRows.Count -gt 0) {
        $nameIdx = $ColumnNames.IndexOf($NameColumn)
        if ($nameIdx -ge 0) { return $dgv.SelectedRows[0].Cells[$nameIdx].Value }
    }
    return $null
}

function Add-LogLine {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TextBox]$TextBox,
        [Parameter(Mandatory)][string]$Message
    )
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "{0}  {1}" -f $ts, $Message

    if ([string]::IsNullOrWhiteSpace($TextBox.Text)) {
        $TextBox.Text = $line
    }
    else {
        $TextBox.AppendText([Environment]::NewLine + $line)
    }

    $TextBox.SelectionStart = $TextBox.TextLength
    $TextBox.ScrollToCaret()
}

function Save-WindowState {
    $statePath = Join-Path $PSScriptRoot "DeploymentHelper.windowstate.json"
    $state = @{
        X         = $form.Location.X
        Y         = $form.Location.Y
        Width     = $form.Size.Width
        Height    = $form.Size.Height
        Maximized = ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Maximized)
    }
    $state | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
}

function Restore-WindowState {
    $statePath = Join-Path $PSScriptRoot "DeploymentHelper.windowstate.json"
    if (-not (Test-Path -LiteralPath $statePath)) { return }

    try {
        $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($state.Maximized) {
            $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
        } else {
            $form.Location = New-Object System.Drawing.Point($state.X, $state.Y)
            $w = [Math]::Max($state.Width, $form.MinimumSize.Width)
            $h = [Math]::Max($state.Height, $form.MinimumSize.Height)
            $form.Size = New-Object System.Drawing.Size($w, $h)
        }
    } catch { }
}

# ---------------------------------------------------------------------------
# Preferences
# ---------------------------------------------------------------------------

function Get-DhPreferences {
    $prefsPath = Join-Path $PSScriptRoot "DeploymentHelper.prefs.json"
    $defaults = @{
        DarkMode          = $false
        SiteCode          = ''
        SMSProvider       = ''
        DeploymentLogPath = ''
    }

    if (Test-Path -LiteralPath $prefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
            if ($null -ne $loaded.DarkMode)          { $defaults.DarkMode          = [bool]$loaded.DarkMode }
            if ($loaded.SiteCode)                    { $defaults.SiteCode          = $loaded.SiteCode }
            if ($loaded.SMSProvider)                  { $defaults.SMSProvider        = $loaded.SMSProvider }
            if ($null -ne $loaded.DeploymentLogPath) { $defaults.DeploymentLogPath = [string]$loaded.DeploymentLogPath }
        } catch { }
    }

    return $defaults
}

function Save-DhPreferences {
    param([hashtable]$Prefs)
    $prefsPath = Join-Path $PSScriptRoot "DeploymentHelper.prefs.json"
    $Prefs | ConvertTo-Json | Set-Content -LiteralPath $prefsPath -Encoding UTF8
}

$script:Prefs = Get-DhPreferences

# ---------------------------------------------------------------------------
# Colors (theme-aware)
# ---------------------------------------------------------------------------

$clrAccent = [System.Drawing.Color]::FromArgb(0, 120, 212)

if ($script:Prefs.DarkMode) {
    $clrFormBg     = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $clrPanelBg    = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $clrHint       = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle   = [System.Drawing.Color]::FromArgb(180, 200, 220)
    $clrGridAlt    = [System.Drawing.Color]::FromArgb(48, 48, 48)
    $clrGridLine   = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $clrDetailBg   = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $clrSepLine    = [System.Drawing.Color]::FromArgb(55, 55, 55)
    $clrInputBdr   = [System.Drawing.Color]::FromArgb(70, 70, 70)
    $clrLogBg      = [System.Drawing.Color]::FromArgb(35, 35, 35)
    $clrLogFg      = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $clrText       = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrGridText   = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrErrText    = [System.Drawing.Color]::FromArgb(255, 100, 100)
    $clrWarnText   = [System.Drawing.Color]::FromArgb(255, 200, 80)
    $clrOkText     = [System.Drawing.Color]::FromArgb(80, 200, 80)
} else {
    $clrFormBg     = [System.Drawing.Color]::FromArgb(245, 246, 248)
    $clrPanelBg    = [System.Drawing.Color]::White
    $clrHint       = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle   = [System.Drawing.Color]::FromArgb(220, 230, 245)
    $clrGridAlt    = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $clrGridLine   = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $clrDetailBg   = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $clrSepLine    = [System.Drawing.Color]::FromArgb(218, 220, 224)
    $clrInputBdr   = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $clrLogBg      = [System.Drawing.Color]::White
    $clrLogFg      = [System.Drawing.Color]::Black
    $clrText       = [System.Drawing.Color]::Black
    $clrGridText   = [System.Drawing.Color]::Black
    $clrErrText    = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $clrWarnText   = [System.Drawing.Color]::FromArgb(180, 120, 0)
    $clrOkText     = [System.Drawing.Color]::FromArgb(34, 139, 34)
}

# Custom dark mode ToolStrip renderer
if ($script:Prefs.DarkMode) {
    if (-not ('DarkToolStripRenderer' -as [type])) {
        $rendererCs = (
            'using System.Drawing;',
            'using System.Windows.Forms;',
            'public class DarkToolStripRenderer : ToolStripProfessionalRenderer {',
            '    private Color _bg;',
            '    public DarkToolStripRenderer(Color bg) : base() { _bg = bg; }',
            '    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }',
            '    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) {',
            '        using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); }',
            '    }',
            '    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) {',
            '        if (e.Item.Selected || e.Item.Pressed) {',
            '            using (var b = new SolidBrush(Color.FromArgb(60, 60, 60)))',
            '            { e.Graphics.FillRectangle(b, new Rectangle(Point.Empty, e.Item.Size)); }',
            '        }',
            '    }',
            '    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e) {',
            '        int y = e.Item.Height / 2;',
            '        using (var p = new Pen(Color.FromArgb(70, 70, 70)))',
            '        { e.Graphics.DrawLine(p, 0, y, e.Item.Width, y); }',
            '    }',
            '    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) {',
            '        using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); }',
            '    }',
            '}'
        ) -join "`r`n"
        Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition $rendererCs
    }
    $script:DarkRenderer = New-Object DarkToolStripRenderer($clrPanelBg)
}

# ---------------------------------------------------------------------------
# Dialogs
# ---------------------------------------------------------------------------

function Show-PreferencesDialog {
    $scriptFile = Join-Path $PSScriptRoot "start-deploymenthelper.ps1"
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Preferences"
    $dlg.Size = New-Object System.Drawing.Size(420, 380)
    $dlg.MinimumSize = $dlg.Size
    $dlg.MaximumSize = $dlg.Size
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.BackColor = $clrFormBg

    # Appearance
    $grpAppearance = New-Object System.Windows.Forms.GroupBox
    $grpAppearance.Text = "Appearance"
    $grpAppearance.SetBounds(16, 12, 372, 60)
    $grpAppearance.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpAppearance.ForeColor = $clrText
    $grpAppearance.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpAppearance.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpAppearance.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpAppearance)

    $chkDark = New-Object System.Windows.Forms.CheckBox
    $chkDark.Text = "Enable dark mode (requires restart)"
    $chkDark.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkDark.AutoSize = $true
    $chkDark.Location = New-Object System.Drawing.Point(14, 24)
    $chkDark.Checked = $script:Prefs.DarkMode
    $chkDark.ForeColor = $clrText
    $chkDark.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $chkDark.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkDark.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
    $grpAppearance.Controls.Add($chkDark)

    # MECM Connection
    $grpConn = New-Object System.Windows.Forms.GroupBox
    $grpConn.Text = "MECM Connection"
    $grpConn.SetBounds(16, 82, 372, 110)
    $grpConn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpConn.ForeColor = $clrText
    $grpConn.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpConn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpConn.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpConn)

    $lblSiteCode = New-Object System.Windows.Forms.Label
    $lblSiteCode.Text = "Site Code:"
    $lblSiteCode.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSiteCode.Location = New-Object System.Drawing.Point(14, 30)
    $lblSiteCode.AutoSize = $true
    $lblSiteCode.ForeColor = $clrText
    $grpConn.Controls.Add($lblSiteCode)

    $txtSiteCodePref = New-Object System.Windows.Forms.TextBox
    $txtSiteCodePref.SetBounds(130, 27, 80, 24)
    $txtSiteCodePref.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtSiteCodePref.MaxLength = 3
    $txtSiteCodePref.Text = $script:Prefs.SiteCode
    $txtSiteCodePref.BackColor = $clrDetailBg
    $txtSiteCodePref.ForeColor = $clrText
    $txtSiteCodePref.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
    $grpConn.Controls.Add($txtSiteCodePref)

    $lblServer = New-Object System.Windows.Forms.Label
    $lblServer.Text = "SMS Provider:"
    $lblServer.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblServer.Location = New-Object System.Drawing.Point(14, 64)
    $lblServer.AutoSize = $true
    $lblServer.ForeColor = $clrText
    $grpConn.Controls.Add($lblServer)

    $txtServerPref = New-Object System.Windows.Forms.TextBox
    $txtServerPref.SetBounds(130, 61, 220, 24)
    $txtServerPref.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtServerPref.Text = $script:Prefs.SMSProvider
    $txtServerPref.BackColor = $clrDetailBg
    $txtServerPref.ForeColor = $clrText
    $txtServerPref.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
    $grpConn.Controls.Add($txtServerPref)

    # Deployment Log
    $grpLog = New-Object System.Windows.Forms.GroupBox
    $grpLog.Text = "Deployment Log"
    $grpLog.SetBounds(16, 202, 372, 70)
    $grpLog.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpLog.ForeColor = $clrText
    $grpLog.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpLog.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpLog.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpLog)

    $lblLogPath = New-Object System.Windows.Forms.Label
    $lblLogPath.Text = "Path:"
    $lblLogPath.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblLogPath.Location = New-Object System.Drawing.Point(14, 30)
    $lblLogPath.AutoSize = $true
    $lblLogPath.ForeColor = $clrText
    $grpLog.Controls.Add($lblLogPath)

    $txtLogPathPref = New-Object System.Windows.Forms.TextBox
    $txtLogPathPref.SetBounds(60, 27, 220, 24)
    $txtLogPathPref.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtLogPathPref.Text = $script:Prefs.DeploymentLogPath
    $txtLogPathPref.BackColor = $clrDetailBg
    $txtLogPathPref.ForeColor = $clrText
    $txtLogPathPref.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
    $grpLog.Controls.Add($txtLogPathPref)

    $btnBrowseLog = New-Object System.Windows.Forms.Button
    $btnBrowseLog.Text = "Browse..."
    $btnBrowseLog.Size = New-Object System.Drawing.Size(72, 24)
    $btnBrowseLog.Location = New-Object System.Drawing.Point(286, 26)
    $btnBrowseLog.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnBrowseLog.FlatAppearance.BorderColor = $clrSepLine
    $btnBrowseLog.ForeColor = $clrText
    $btnBrowseLog.BackColor = $clrFormBg
    $btnBrowseLog.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "Select deployment log folder"
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtLogPathPref.Text = $fbd.SelectedPath
        }
    })
    $grpLog.Controls.Add($btnBrowseLog)

    # OK / Cancel
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Size = New-Object System.Drawing.Size(90, 32)
    $btnOK.Location = New-Object System.Drawing.Point(208, 290)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    Set-ModernButtonStyle -Button $btnOK -BackColor $clrAccent
    $dlg.Controls.Add($btnOK)
    $dlg.AcceptButton = $btnOK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnCancel.Location = New-Object System.Drawing.Point(306, 290)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderColor = $clrSepLine
    $btnCancel.ForeColor = $clrText
    $btnCancel.BackColor = $clrFormBg
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel

    if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $darkChanged = ($chkDark.Checked -ne $script:Prefs.DarkMode)
        $script:Prefs.DarkMode          = $chkDark.Checked
        $script:Prefs.SiteCode          = $txtSiteCodePref.Text.Trim().ToUpper()
        $script:Prefs.SMSProvider       = $txtServerPref.Text.Trim()
        $script:Prefs.DeploymentLogPath = $txtLogPathPref.Text.Trim()
        Save-DhPreferences -Prefs $script:Prefs

        # Update connection bar labels
        $lblSiteVal.Text   = if ($script:Prefs.SiteCode)    { $script:Prefs.SiteCode }    else { '(not set)' }
        $lblServerVal.Text = if ($script:Prefs.SMSProvider)  { $script:Prefs.SMSProvider }  else { '(not set)' }

        if ($darkChanged) {
            $restart = [System.Windows.Forms.MessageBox]::Show(
                "Theme change requires a restart. Restart now?",
                "Restart Required",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($restart -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process powershell -ArgumentList @('-ExecutionPolicy', 'Bypass', '-File', "`"$scriptFile`"")
                $form.Close()
            }
        }
    }

    $dlg.Dispose()
}

function Show-AboutDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "About Deployment Helper"
    $dlg.Size = New-Object System.Drawing.Size(460, 320)
    $dlg.MinimumSize = $dlg.Size
    $dlg.MaximumSize = $dlg.Size
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $dlg.BackColor = $clrFormBg

    $lblAboutTitle = New-Object System.Windows.Forms.Label
    $lblAboutTitle.Text = "Deployment Helper"
    $lblAboutTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblAboutTitle.ForeColor = $clrAccent
    $lblAboutTitle.AutoSize = $true
    $lblAboutTitle.BackColor = $clrFormBg
    $lblAboutTitle.Location = New-Object System.Drawing.Point(120, 30)
    $dlg.Controls.Add($lblAboutTitle)

    $lblVersion = New-Object System.Windows.Forms.Label
    $lblVersion.Text = "Deployment Helper v1.3.0"
    $lblVersion.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblVersion.ForeColor = $clrText
    $lblVersion.AutoSize = $true
    $lblVersion.BackColor = $clrFormBg
    $lblVersion.Location = New-Object System.Drawing.Point(130, 60)
    $dlg.Controls.Add($lblVersion)

    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = ("Safe, fast MECM application deployment with pre-execution validation," +
        " safety guardrails, deployment templates, and immutable audit logging." +
        " Reduces deployment from 10-15 minutes to 15-30 seconds.")
    $lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblDesc.ForeColor = $clrText
    $lblDesc.SetBounds(30, 100, 390, 80)
    $lblDesc.BackColor = $clrFormBg
    $lblDesc.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
    $dlg.Controls.Add($lblDesc)

    $lblCopyright = New-Object System.Windows.Forms.Label
    $lblCopyright.Text = "(c) 2026 - All rights reserved"
    $lblCopyright.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)
    $lblCopyright.ForeColor = $clrHint
    $lblCopyright.AutoSize = $true
    $lblCopyright.BackColor = $clrFormBg
    $lblCopyright.Location = New-Object System.Drawing.Point(142, 200)
    $dlg.Controls.Add($lblCopyright)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "OK"
    $btnClose.Size = New-Object System.Drawing.Size(90, 32)
    $btnClose.Location = New-Object System.Drawing.Point(175, 240)
    $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::OK
    Set-ModernButtonStyle -Button $btnClose -BackColor $clrAccent
    $dlg.Controls.Add($btnClose)
    $dlg.AcceptButton = $btnClose

    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
}

# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Deployment Helper"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(900, 988)
$form.MinimumSize = New-Object System.Drawing.Size(780, 900)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.BackColor = $clrFormBg
$form.Icon = [System.Drawing.SystemIcons]::Application

# ---------------------------------------------------------------------------
# Menu bar
# ---------------------------------------------------------------------------

$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Dock = [System.Windows.Forms.DockStyle]::Top
$menuStrip.BackColor = $clrPanelBg
$menuStrip.ForeColor = $clrText
if ($script:DarkRenderer) {
    $menuStrip.Renderer = $script:DarkRenderer
} else {
    $menuStrip.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::System
}
$menuStrip.Padding = New-Object System.Windows.Forms.Padding(4, 2, 0, 0)

# File menu
$mnuFile = New-Object System.Windows.Forms.ToolStripMenuItem("&File")
$mnuFilePrefs = New-Object System.Windows.Forms.ToolStripMenuItem("&Preferences...")
$mnuFilePrefs.Add_Click({ Show-PreferencesDialog })
$mnuFileSep = New-Object System.Windows.Forms.ToolStripSeparator
$mnuFileExit = New-Object System.Windows.Forms.ToolStripMenuItem("E&xit")
$mnuFileExit.Add_Click({ $form.Close() })
[void]$mnuFile.DropDownItems.Add($mnuFilePrefs)
[void]$mnuFile.DropDownItems.Add($mnuFileSep)
[void]$mnuFile.DropDownItems.Add($mnuFileExit)

# Help menu
$mnuHelp = New-Object System.Windows.Forms.ToolStripMenuItem("&Help")
$mnuHelpAbout = New-Object System.Windows.Forms.ToolStripMenuItem("&About...")
$mnuHelpAbout.Add_Click({ Show-AboutDialog })
[void]$mnuHelp.DropDownItems.Add($mnuHelpAbout)

[void]$menuStrip.Items.Add($mnuFile)
[void]$menuStrip.Items.Add($mnuHelp)
$form.MainMenuStrip = $menuStrip

# ---------------------------------------------------------------------------
# StatusStrip (Dock:Bottom - add FIRST so it stays at very bottom)
# ---------------------------------------------------------------------------

$status = New-Object System.Windows.Forms.StatusStrip
$status.BackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(45, 45, 45) } else { [System.Drawing.Color]::FromArgb(240, 240, 240) }
$status.ForeColor = $clrText
$status.Dock = [System.Windows.Forms.DockStyle]::Bottom
if ($script:DarkRenderer) {
    $status.Renderer = $script:DarkRenderer
} else {
    $status.RenderMode = [System.Windows.Forms.ToolStripRenderMode]::System
}
$status.SizingGrip = $false
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Configure site in File > Preferences, then click Connect."
$statusLabel.ForeColor = $clrText
$status.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($status)

# ---------------------------------------------------------------------------
# Log console panel (Dock:Bottom)
# ---------------------------------------------------------------------------

$pnlLog = New-Object System.Windows.Forms.Panel
$pnlLog.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlLog.Height = 95
$pnlLog.Padding = New-Object System.Windows.Forms.Padding(12, 4, 12, 6)
$pnlLog.BackColor = $clrFormBg
$form.Controls.Add($pnlLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = if ($script:Prefs.DarkMode) { [System.Windows.Forms.ScrollBars]::None } else { [System.Windows.Forms.ScrollBars]::Vertical }
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLog.BackColor = $clrLogBg
$txtLog.ForeColor = $clrLogFg
$txtLog.WordWrap = $true
$txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$pnlLog.Controls.Add($txtLog)

# ---------------------------------------------------------------------------
# Button panel (Dock:Bottom)
# ---------------------------------------------------------------------------

$pnlButtons = New-Object System.Windows.Forms.Panel
$pnlButtons.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlButtons.Height = 56
$pnlButtons.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 4)
$pnlButtons.BackColor = $clrFormBg
$form.Controls.Add($pnlButtons)

$pnlSepButtons = New-Object System.Windows.Forms.Panel
$pnlSepButtons.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSepButtons.Height = 1
$pnlSepButtons.BackColor = $clrSepLine
$pnlButtons.Controls.Add($pnlSepButtons)

$flowButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$flowButtons.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowButtons.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowButtons.WrapContents = $false
$flowButtons.BackColor = $clrFormBg
$pnlButtons.Controls.Add($flowButtons)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = "Export History CSV"
$btnExportCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$btnExportCsv.Size = New-Object System.Drawing.Size(160, 38)
$btnExportCsv.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
Set-ModernButtonStyle -Button $btnExportCsv -BackColor ([System.Drawing.Color]::FromArgb(34, 139, 34))
$flowButtons.Controls.Add($btnExportCsv)

$btnExportHtml = New-Object System.Windows.Forms.Button
$btnExportHtml.Text = "Export History HTML"
$btnExportHtml.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$btnExportHtml.Size = New-Object System.Drawing.Size(170, 38)
$btnExportHtml.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
Set-ModernButtonStyle -Button $btnExportHtml -BackColor ([System.Drawing.Color]::FromArgb(34, 139, 34))
$flowButtons.Controls.Add($btnExportHtml)

# ---------------------------------------------------------------------------
# Header panel (Dock:Top)
# ---------------------------------------------------------------------------

$pnlHeader = New-Object System.Windows.Forms.Panel
$pnlHeader.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlHeader.Height = 60
$pnlHeader.BackColor = $clrAccent
$pnlHeader.Padding = New-Object System.Windows.Forms.Padding(16, 0, 16, 0)
$form.Controls.Add($pnlHeader)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Deployment Helper"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 17, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.AutoSize = $true
$lblTitle.BackColor = [System.Drawing.Color]::Transparent
$lblTitle.Location = New-Object System.Drawing.Point(16, 8)
$pnlHeader.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = "Safe MECM Application Deployment"
$lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSubtitle.ForeColor = $clrSubtitle
$lblSubtitle.AutoSize = $true
$lblSubtitle.BackColor = [System.Drawing.Color]::Transparent
$lblSubtitle.Location = New-Object System.Drawing.Point(18, 36)
$pnlHeader.Controls.Add($lblSubtitle)

# ---------------------------------------------------------------------------
# Connection bar (Dock:Top)
# ---------------------------------------------------------------------------

$pnlConnBar = New-Object System.Windows.Forms.Panel
$pnlConnBar.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlConnBar.Height = 36
$pnlConnBar.BackColor = $clrPanelBg
$pnlConnBar.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6)
$form.Controls.Add($pnlConnBar)

$flowConn = New-Object System.Windows.Forms.FlowLayoutPanel
$flowConn.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowConn.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$flowConn.WrapContents = $false
$flowConn.BackColor = $clrPanelBg
$pnlConnBar.Controls.Add($flowConn)

$lblSiteLabel = New-Object System.Windows.Forms.Label
$lblSiteLabel.Text = "Site:"
$lblSiteLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblSiteLabel.AutoSize = $true
$lblSiteLabel.Margin = New-Object System.Windows.Forms.Padding(0, 3, 2, 0)
$lblSiteLabel.ForeColor = $clrText
$lblSiteLabel.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSiteLabel)

$lblSiteVal = New-Object System.Windows.Forms.Label
$lblSiteVal.Text = if ($script:Prefs.SiteCode) { $script:Prefs.SiteCode } else { '(not set)' }
$lblSiteVal.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblSiteVal.AutoSize = $true
$lblSiteVal.Margin = New-Object System.Windows.Forms.Padding(0, 3, 16, 0)
$lblSiteVal.ForeColor = if ($script:Prefs.SiteCode) { $clrAccent } else { $clrHint }
$lblSiteVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSiteVal)

$lblServerLabel = New-Object System.Windows.Forms.Label
$lblServerLabel.Text = "Server:"
$lblServerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblServerLabel.AutoSize = $true
$lblServerLabel.Margin = New-Object System.Windows.Forms.Padding(0, 3, 2, 0)
$lblServerLabel.ForeColor = $clrText
$lblServerLabel.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblServerLabel)

$lblServerVal = New-Object System.Windows.Forms.Label
$lblServerVal.Text = if ($script:Prefs.SMSProvider) { $script:Prefs.SMSProvider } else { '(not set)' }
$lblServerVal.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblServerVal.AutoSize = $true
$lblServerVal.Margin = New-Object System.Windows.Forms.Padding(0, 3, 16, 0)
$lblServerVal.ForeColor = if ($script:Prefs.SMSProvider) { $clrAccent } else { $clrHint }
$lblServerVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblServerVal)

$lblConnStatus = New-Object System.Windows.Forms.Label
$lblConnStatus.Text = "Disconnected"
$lblConnStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblConnStatus.AutoSize = $true
$lblConnStatus.Margin = New-Object System.Windows.Forms.Padding(0, 3, 20, 0)
$lblConnStatus.ForeColor = $clrHint
$lblConnStatus.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblConnStatus)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect"
$btnConnect.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnConnect.Size = New-Object System.Drawing.Size(90, 24)
$btnConnect.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
Set-ModernButtonStyle -Button $btnConnect -BackColor $clrAccent
$flowConn.Controls.Add($btnConnect)

# Separator below connection bar
$pnlSep1 = New-Object System.Windows.Forms.Panel
$pnlSep1.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep1.Height = 1
$pnlSep1.BackColor = $clrSepLine
$form.Controls.Add($pnlSep1)

# ---------------------------------------------------------------------------
# Deployment form panel (Dock:Top)
# ---------------------------------------------------------------------------

$pnlForm = New-Object System.Windows.Forms.Panel
$pnlForm.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlForm.Height = 480
$pnlForm.BackColor = $clrPanelBg
$form.Controls.Add($pnlForm)

# Load templates
$script:Templates = Get-DeploymentTemplates -TemplatePath (Join-Path $PSScriptRoot "Templates")

# Row 1: Deployment Type
$lblDeployType = New-Object System.Windows.Forms.Label
$lblDeployType.Text = "Type:"
$lblDeployType.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblDeployType.ForeColor = $clrText
$lblDeployType.Location = New-Object System.Drawing.Point(14, 14)
$lblDeployType.AutoSize = $true
$pnlForm.Controls.Add($lblDeployType)

$cboDeployType = New-Object System.Windows.Forms.ComboBox
$cboDeployType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboDeployType.SetBounds(160, 11, 220, 24)
$cboDeployType.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboDeployType.BackColor = $clrDetailBg
$cboDeployType.ForeColor = $clrText
$cboDeployType.FlatStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.FlatStyle]::Flat } else { [System.Windows.Forms.FlatStyle]::Standard }
[void]$cboDeployType.Items.AddRange(@('Application', 'Software Update Group'))
$cboDeployType.SelectedIndex = 0
$pnlForm.Controls.Add($cboDeployType)

# Row 3: Application / SUG Name
$lblAppName = New-Object System.Windows.Forms.Label
$lblAppName.Text = "Application:"
$lblAppName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblAppName.ForeColor = $clrText
$lblAppName.Location = New-Object System.Drawing.Point(14, 46)
$lblAppName.AutoSize = $true
$pnlForm.Controls.Add($lblAppName)

$txtAppName = New-Object System.Windows.Forms.TextBox
$txtAppName.SetBounds(160, 43, 250, 24)
$txtAppName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtAppName.BackColor = $clrDetailBg
$txtAppName.ForeColor = $clrText
$txtAppName.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
$pnlForm.Controls.Add($txtAppName)

$btnBrowseApp = New-Object System.Windows.Forms.Button
$btnBrowseApp.Text = "Browse"
$btnBrowseApp.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$btnBrowseApp.Size = New-Object System.Drawing.Size(56, 24)
$btnBrowseApp.Location = New-Object System.Drawing.Point(414, 43)
$btnBrowseApp.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBrowseApp.FlatAppearance.BorderColor = $clrSepLine
$btnBrowseApp.ForeColor = $clrText
$btnBrowseApp.BackColor = $clrFormBg
$btnBrowseApp.Cursor = [System.Windows.Forms.Cursors]::Hand
$pnlForm.Controls.Add($btnBrowseApp)

# Row 4: Collection
$lblCollName = New-Object System.Windows.Forms.Label
$lblCollName.Text = "Collection:"
$lblCollName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblCollName.ForeColor = $clrText
$lblCollName.Location = New-Object System.Drawing.Point(14, 78)
$lblCollName.AutoSize = $true
$pnlForm.Controls.Add($lblCollName)

$txtCollName = New-Object System.Windows.Forms.TextBox
$txtCollName.SetBounds(160, 75, 250, 24)
$txtCollName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtCollName.BackColor = $clrDetailBg
$txtCollName.ForeColor = $clrText
$txtCollName.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
$pnlForm.Controls.Add($txtCollName)

$btnBrowseColl = New-Object System.Windows.Forms.Button
$btnBrowseColl.Text = "Browse"
$btnBrowseColl.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$btnBrowseColl.Size = New-Object System.Drawing.Size(56, 24)
$btnBrowseColl.Location = New-Object System.Drawing.Point(414, 75)
$btnBrowseColl.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBrowseColl.FlatAppearance.BorderColor = $clrSepLine
$btnBrowseColl.ForeColor = $clrText
$btnBrowseColl.BackColor = $clrFormBg
$btnBrowseColl.Cursor = [System.Windows.Forms.Cursors]::Hand
$pnlForm.Controls.Add($btnBrowseColl)

# Row 5: Distribute To (DP Groups)
$lblDPGroups = New-Object System.Windows.Forms.Label
$lblDPGroups.Text = "Distribute to:"
$lblDPGroups.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblDPGroups.ForeColor = $clrText
$lblDPGroups.Location = New-Object System.Drawing.Point(14, 110)
$lblDPGroups.AutoSize = $true
$pnlForm.Controls.Add($lblDPGroups)

$clbDPGroups = New-Object System.Windows.Forms.CheckedListBox
$clbDPGroups.SetBounds(160, 107, 300, 56)
$clbDPGroups.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$clbDPGroups.BackColor = $clrDetailBg
$clbDPGroups.ForeColor = $clrText
$clbDPGroups.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
$clbDPGroups.CheckOnClick = $true
$clbDPGroups.Enabled = $false
$pnlForm.Controls.Add($clbDPGroups)

# Row 6: Template
$lblTemplate = New-Object System.Windows.Forms.Label
$lblTemplate.Text = "Template:"
$lblTemplate.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblTemplate.ForeColor = $clrText
$lblTemplate.Location = New-Object System.Drawing.Point(14, 178)
$lblTemplate.AutoSize = $true
$pnlForm.Controls.Add($lblTemplate)

$cboTemplate = New-Object System.Windows.Forms.ComboBox
$cboTemplate.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboTemplate.SetBounds(160, 175, 220, 24)
$cboTemplate.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboTemplate.BackColor = $clrDetailBg
$cboTemplate.ForeColor = $clrText
$cboTemplate.FlatStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.FlatStyle]::Flat } else { [System.Windows.Forms.FlatStyle]::Standard }
[void]$cboTemplate.Items.Add("(None)")
foreach ($tmpl in $script:Templates) { [void]$cboTemplate.Items.Add($tmpl.Name) }
$cboTemplate.SelectedIndex = 0
$pnlForm.Controls.Add($cboTemplate)

# Row 7: Purpose
$lblPurpose = New-Object System.Windows.Forms.Label
$lblPurpose.Text = "Purpose:"
$lblPurpose.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblPurpose.ForeColor = $clrText
$lblPurpose.Location = New-Object System.Drawing.Point(14, 210)
$lblPurpose.AutoSize = $true
$pnlForm.Controls.Add($lblPurpose)

$radAvailable = New-Object System.Windows.Forms.RadioButton
$radAvailable.Text = "Available"
$radAvailable.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radAvailable.ForeColor = $clrText
$radAvailable.BackColor = $clrPanelBg
$radAvailable.Location = New-Object System.Drawing.Point(160, 208)
$radAvailable.AutoSize = $true
$radAvailable.Checked = $true
if ($script:Prefs.DarkMode) { $radAvailable.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $radAvailable.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($radAvailable)

$radRequired = New-Object System.Windows.Forms.RadioButton
$radRequired.Text = "Required"
$radRequired.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radRequired.ForeColor = $clrText
$radRequired.BackColor = $clrPanelBg
$radRequired.Location = New-Object System.Drawing.Point(270, 208)
$radRequired.AutoSize = $true
if ($script:Prefs.DarkMode) { $radRequired.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $radRequired.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($radRequired)

# Row 7: Available date
$lblAvailable = New-Object System.Windows.Forms.Label
$lblAvailable.Text = "Available:"
$lblAvailable.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblAvailable.ForeColor = $clrText
$lblAvailable.Location = New-Object System.Drawing.Point(14, 242)
$lblAvailable.AutoSize = $true
$pnlForm.Controls.Add($lblAvailable)

$dtpAvailable = New-Object System.Windows.Forms.DateTimePicker
$dtpAvailable.SetBounds(160, 239, 200, 24)
$dtpAvailable.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$dtpAvailable.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpAvailable.CustomFormat = "yyyy-MM-dd HH:mm"
$dtpAvailable.Value = Get-Date
$pnlForm.Controls.Add($dtpAvailable)

# Row 8: Deadline date
$lblDeadline = New-Object System.Windows.Forms.Label
$lblDeadline.Text = "Deadline:"
$lblDeadline.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblDeadline.ForeColor = $clrText
$lblDeadline.Location = New-Object System.Drawing.Point(14, 274)
$lblDeadline.AutoSize = $true
$pnlForm.Controls.Add($lblDeadline)

$dtpDeadline = New-Object System.Windows.Forms.DateTimePicker
$dtpDeadline.SetBounds(160, 271, 200, 24)
$dtpDeadline.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$dtpDeadline.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpDeadline.CustomFormat = "yyyy-MM-dd HH:mm"
$dtpDeadline.Value = (Get-Date).AddHours(24)
$dtpDeadline.Enabled = $false
$pnlForm.Controls.Add($dtpDeadline)

# Row 9: Notification
$lblNotification = New-Object System.Windows.Forms.Label
$lblNotification.Text = "Notification:"
$lblNotification.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblNotification.ForeColor = $clrText
$lblNotification.Location = New-Object System.Drawing.Point(14, 306)
$lblNotification.AutoSize = $true
$pnlForm.Controls.Add($lblNotification)

$cboNotification = New-Object System.Windows.Forms.ComboBox
$cboNotification.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboNotification.SetBounds(160, 303, 260, 24)
$cboNotification.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboNotification.BackColor = $clrDetailBg
$cboNotification.ForeColor = $clrText
$cboNotification.FlatStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.FlatStyle]::Flat } else { [System.Windows.Forms.FlatStyle]::Standard }
[void]$cboNotification.Items.AddRange(@('Display All Notifications', 'Display in Software Center Only', 'Hide All Notifications'))
$cboNotification.SelectedIndex = 0
$pnlForm.Controls.Add($cboNotification)

# Row 10: Time basis
$lblTimeBasis = New-Object System.Windows.Forms.Label
$lblTimeBasis.Text = "Time basis:"
$lblTimeBasis.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblTimeBasis.ForeColor = $clrText
$lblTimeBasis.Location = New-Object System.Drawing.Point(14, 334)
$lblTimeBasis.AutoSize = $true
$pnlForm.Controls.Add($lblTimeBasis)

$cboTimeBasis = New-Object System.Windows.Forms.ComboBox
$cboTimeBasis.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboTimeBasis.SetBounds(160, 331, 180, 24)
$cboTimeBasis.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboTimeBasis.BackColor = $clrDetailBg
$cboTimeBasis.ForeColor = $clrText
$cboTimeBasis.FlatStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.FlatStyle]::Flat } else { [System.Windows.Forms.FlatStyle]::Standard }
[void]$cboTimeBasis.Items.AddRange(@('Client Local Time', 'UTC'))
$cboTimeBasis.SelectedIndex = 0
$pnlForm.Controls.Add($cboTimeBasis)

# Row 11: Maintenance window overrides
$chkOverrideMW = New-Object System.Windows.Forms.CheckBox
$chkOverrideMW.Text = "Allow outside maintenance window"
$chkOverrideMW.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkOverrideMW.ForeColor = $clrText
$chkOverrideMW.BackColor = $clrPanelBg
$chkOverrideMW.Location = New-Object System.Drawing.Point(160, 362)
$chkOverrideMW.AutoSize = $true
if ($script:Prefs.DarkMode) { $chkOverrideMW.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkOverrideMW.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($chkOverrideMW)

$chkRebootOutside = New-Object System.Windows.Forms.CheckBox
$chkRebootOutside.Text = "Reboot outside maintenance window"
$chkRebootOutside.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkRebootOutside.ForeColor = $clrText
$chkRebootOutside.BackColor = $clrPanelBg
$chkRebootOutside.Location = New-Object System.Drawing.Point(160, 384)
$chkRebootOutside.AutoSize = $true
if ($script:Prefs.DarkMode) { $chkRebootOutside.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkRebootOutside.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($chkRebootOutside)

# Row 12: Metered connection
$chkMetered = New-Object System.Windows.Forms.CheckBox
$chkMetered.Text = "Allow download past deadline (metered connections)"
$chkMetered.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkMetered.ForeColor = $clrText
$chkMetered.BackColor = $clrPanelBg
$chkMetered.Location = New-Object System.Drawing.Point(160, 406)
$chkMetered.AutoSize = $true
if ($script:Prefs.DarkMode) { $chkMetered.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkMetered.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($chkMetered)

# Row 13: SUG-specific download settings
$chkBoundaryFallback = New-Object System.Windows.Forms.CheckBox
$chkBoundaryFallback.Text = "Allow download from default site boundary group"
$chkBoundaryFallback.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkBoundaryFallback.ForeColor = $clrText
$chkBoundaryFallback.BackColor = $clrPanelBg
$chkBoundaryFallback.Location = New-Object System.Drawing.Point(160, 428)
$chkBoundaryFallback.AutoSize = $true
$chkBoundaryFallback.Checked = $true
$chkBoundaryFallback.Visible = $false
if ($script:Prefs.DarkMode) { $chkBoundaryFallback.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkBoundaryFallback.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($chkBoundaryFallback)

$chkMicrosoftUpdate = New-Object System.Windows.Forms.CheckBox
$chkMicrosoftUpdate.Text = "Allow download from Microsoft Update"
$chkMicrosoftUpdate.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkMicrosoftUpdate.ForeColor = $clrText
$chkMicrosoftUpdate.BackColor = $clrPanelBg
$chkMicrosoftUpdate.Location = New-Object System.Drawing.Point(160, 450)
$chkMicrosoftUpdate.AutoSize = $true
$chkMicrosoftUpdate.Visible = $false
if ($script:Prefs.DarkMode) { $chkMicrosoftUpdate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkMicrosoftUpdate.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($chkMicrosoftUpdate)

$chkPostRebootScan = New-Object System.Windows.Forms.CheckBox
$chkPostRebootScan.Text = "Require post-reboot full scan"
$chkPostRebootScan.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkPostRebootScan.ForeColor = $clrText
$chkPostRebootScan.BackColor = $clrPanelBg
$chkPostRebootScan.Location = New-Object System.Drawing.Point(160, 472)
$chkPostRebootScan.AutoSize = $true
$chkPostRebootScan.Checked = $true
$chkPostRebootScan.Visible = $false
if ($script:Prefs.DarkMode) { $chkPostRebootScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkPostRebootScan.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($chkPostRebootScan)

# Row 14: Validate + Deploy + Save Template buttons
$btnValidate = New-Object System.Windows.Forms.Button
$btnValidate.Text = "Validate"
$btnValidate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnValidate.Size = New-Object System.Drawing.Size(120, 32)
$btnValidate.Location = New-Object System.Drawing.Point(160, 440)
Set-ModernButtonStyle -Button $btnValidate -BackColor $clrAccent
$pnlForm.Controls.Add($btnValidate)

$btnDeploy = New-Object System.Windows.Forms.Button
$btnDeploy.Text = "Deploy"
$btnDeploy.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnDeploy.Size = New-Object System.Drawing.Size(120, 32)
$btnDeploy.Location = New-Object System.Drawing.Point(290, 440)
$btnDeploy.Enabled = $false
Set-ModernButtonStyle -Button $btnDeploy -BackColor ([System.Drawing.Color]::FromArgb(34, 139, 34))
$pnlForm.Controls.Add($btnDeploy)

$btnSaveTemplate = New-Object System.Windows.Forms.Button
$btnSaveTemplate.Text = "Save Template"
$btnSaveTemplate.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSaveTemplate.Size = New-Object System.Drawing.Size(130, 32)
$btnSaveTemplate.Location = New-Object System.Drawing.Point(420, 440)
$btnSaveTemplate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnSaveTemplate.FlatAppearance.BorderColor = $clrSepLine
$btnSaveTemplate.ForeColor = $clrText
$btnSaveTemplate.BackColor = $clrFormBg
$btnSaveTemplate.Cursor = [System.Windows.Forms.Cursors]::Hand
$pnlForm.Controls.Add($btnSaveTemplate)

# Separator below deployment form
$pnlSep2 = New-Object System.Windows.Forms.Panel
$pnlSep2.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlSep2.Height = 1
$pnlSep2.BackColor = $clrSepLine
$form.Controls.Add($pnlSep2)

# ---------------------------------------------------------------------------
# Validation results panel (Dock:Fill)
# ---------------------------------------------------------------------------

$pnlValidation = New-Object System.Windows.Forms.Panel
$pnlValidation.Dock = [System.Windows.Forms.DockStyle]::Fill
$pnlValidation.BackColor = $clrPanelBg
$pnlValidation.Padding = New-Object System.Windows.Forms.Padding(16, 8, 16, 8)
$form.Controls.Add($pnlValidation)

$rtbValidation = New-Object System.Windows.Forms.RichTextBox
$rtbValidation.Dock = [System.Windows.Forms.DockStyle]::Fill
$rtbValidation.ReadOnly = $true
$rtbValidation.Font = New-Object System.Drawing.Font("Consolas", 9.5)
$rtbValidation.BackColor = $clrDetailBg
$rtbValidation.ForeColor = $clrText
$rtbValidation.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$rtbValidation.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$pnlValidation.Controls.Add($rtbValidation)

# ---------------------------------------------------------------------------
# Dock Z-order finalization
# ---------------------------------------------------------------------------

$form.Controls.Add($menuStrip)
$menuStrip.SendToBack()

# BringToFront processes innermost-last. First call = outermost (top edge).
# Visual order top-to-bottom: Header -> ConnBar -> Sep1 -> Form -> Sep2
$pnlHeader.BringToFront()
$pnlConnBar.BringToFront()
$pnlSep1.BringToFront()
$pnlForm.BringToFront()
$pnlSep2.BringToFront()

# Fill panel must BringToFront last
$pnlValidation.BringToFront()

# ---------------------------------------------------------------------------
# Validation result helper
# ---------------------------------------------------------------------------

function Add-ValidationLine {
    param(
        [string]$Icon,
        [string]$Message,
        [System.Drawing.Color]$Color
    )
    $rtbValidation.SelectionStart = $rtbValidation.TextLength
    $rtbValidation.SelectionLength = 0
    $rtbValidation.SelectionColor = $Color
    $line = "  $Icon  $Message"
    if ($rtbValidation.TextLength -gt 0) { $line = [Environment]::NewLine + $line }
    $rtbValidation.AppendText($line)
    $rtbValidation.ScrollToCaret()
}

# ---------------------------------------------------------------------------
# Event: Required/Available toggle
# ---------------------------------------------------------------------------

$radRequired.Add_CheckedChanged({
    $dtpDeadline.Enabled = $radRequired.Checked
    # Auto-check metered for Required deployments
    if ($radRequired.Checked) { $chkMetered.Checked = $true }
})

# ---------------------------------------------------------------------------
# Event: Deployment type change
# ---------------------------------------------------------------------------

$cboDeployType.Add_SelectedIndexChanged({
    $isSUG = ($cboDeployType.SelectedIndex -eq 1)
    if ($isSUG) {
        $lblAppName.Text = "SUG Name:"
        $pnlForm.Height = 550
        $btnValidate.Location = New-Object System.Drawing.Point(160, 508)
        $btnDeploy.Location = New-Object System.Drawing.Point(290, 508)
        $btnSaveTemplate.Location = New-Object System.Drawing.Point(420, 508)
    } else {
        $lblAppName.Text = "Application:"
        $pnlForm.Height = 480
        $btnValidate.Location = New-Object System.Drawing.Point(160, 440)
        $btnDeploy.Location = New-Object System.Drawing.Point(290, 440)
        $btnSaveTemplate.Location = New-Object System.Drawing.Point(420, 440)
    }
    $chkBoundaryFallback.Visible = $isSUG
    $chkMicrosoftUpdate.Visible  = $isSUG
    $chkPostRebootScan.Visible   = $isSUG
})

# ---------------------------------------------------------------------------
# Event: Template selection
# ---------------------------------------------------------------------------

$cboTemplate.Add_SelectedIndexChanged({
    if ($cboTemplate.SelectedIndex -le 0) { return }
    $tmpl = $script:Templates[$cboTemplate.SelectedIndex - 1]
    if ($tmpl.DeployPurpose -eq 'Required') { $radRequired.Checked = $true } else { $radAvailable.Checked = $true }
    switch ($tmpl.UserNotification) {
        'DisplayAll'               { $cboNotification.SelectedIndex = 0 }
        'DisplaySoftwareCenterOnly' { $cboNotification.SelectedIndex = 1 }
        'HideAll'                  { $cboNotification.SelectedIndex = 2 }
    }
    $chkOverrideMW.Checked = [bool]$tmpl.OverrideServiceWindow
    $chkRebootOutside.Checked = [bool]$tmpl.RebootOutsideServiceWindow
    if ($null -ne $tmpl.AllowMeteredConnection) { $chkMetered.Checked = [bool]$tmpl.AllowMeteredConnection }
    if ($tmpl.TimeBasedOn -eq 'Utc' -or $tmpl.TimeBasedOn -eq 'UTC') { $cboTimeBasis.SelectedIndex = 1 } else { $cboTimeBasis.SelectedIndex = 0 }
    if ($null -ne $tmpl.AllowBoundaryFallback) { $chkBoundaryFallback.Checked = [bool]$tmpl.AllowBoundaryFallback }
    if ($null -ne $tmpl.AllowMicrosoftUpdate) { $chkMicrosoftUpdate.Checked = [bool]$tmpl.AllowMicrosoftUpdate }
    if ($null -ne $tmpl.RequirePostRebootFullScan) { $chkPostRebootScan.Checked = [bool]$tmpl.RequirePostRebootFullScan }
    if ($tmpl.DefaultDeadlineOffsetHours -and $tmpl.DefaultDeadlineOffsetHours -gt 0) {
        $dtpDeadline.Value = (Get-Date).AddHours($tmpl.DefaultDeadlineOffsetHours)
    }
})

# ---------------------------------------------------------------------------
# Event: Save Template
# ---------------------------------------------------------------------------

$btnSaveTemplate.Add_Click({
    $inputDlg = New-Object System.Windows.Forms.Form
    $inputDlg.Text = "Save Template"
    $inputDlg.Size = New-Object System.Drawing.Size(360, 160)
    $inputDlg.MinimumSize = $inputDlg.Size
    $inputDlg.MaximumSize = $inputDlg.Size
    $inputDlg.StartPosition = "CenterParent"
    $inputDlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $inputDlg.MaximizeBox = $false
    $inputDlg.MinimizeBox = $false
    $inputDlg.ShowInTaskbar = $false
    $inputDlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $inputDlg.BackColor = $clrFormBg

    $lblTmplName = New-Object System.Windows.Forms.Label
    $lblTmplName.Text = "Template name:"
    $lblTmplName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblTmplName.ForeColor = $clrText
    $lblTmplName.Location = New-Object System.Drawing.Point(16, 20)
    $lblTmplName.AutoSize = $true
    $inputDlg.Controls.Add($lblTmplName)

    $txtTmplName = New-Object System.Windows.Forms.TextBox
    $txtTmplName.SetBounds(130, 17, 200, 24)
    $txtTmplName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $txtTmplName.BackColor = $clrDetailBg
    $txtTmplName.ForeColor = $clrText
    $inputDlg.Controls.Add($txtTmplName)

    $btnTmplOK = New-Object System.Windows.Forms.Button
    $btnTmplOK.Text = "Save"
    $btnTmplOK.Size = New-Object System.Drawing.Size(90, 32)
    $btnTmplOK.Location = New-Object System.Drawing.Point(150, 70)
    $btnTmplOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    Set-ModernButtonStyle -Button $btnTmplOK -BackColor $clrAccent
    $inputDlg.Controls.Add($btnTmplOK)
    $inputDlg.AcceptButton = $btnTmplOK

    $btnTmplCancel = New-Object System.Windows.Forms.Button
    $btnTmplCancel.Text = "Cancel"
    $btnTmplCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnTmplCancel.Location = New-Object System.Drawing.Point(248, 70)
    $btnTmplCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $btnTmplCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnTmplCancel.FlatAppearance.BorderColor = $clrSepLine
    $btnTmplCancel.ForeColor = $clrText
    $btnTmplCancel.BackColor = $clrFormBg
    $inputDlg.Controls.Add($btnTmplCancel)
    $inputDlg.CancelButton = $btnTmplCancel

    if ($inputDlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $tmplName = $txtTmplName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($tmplName)) {
            [System.Windows.Forms.MessageBox]::Show("Enter a template name.", "Missing Name", "OK", "Warning") | Out-Null
            $inputDlg.Dispose()
            return
        }

        $notifMap = @('DisplayAll', 'DisplaySoftwareCenterOnly', 'HideAll')
        $purpose = if ($radRequired.Checked) { 'Required' } else { 'Available' }

        $deadlineOffset = 0
        if ($radRequired.Checked) {
            $diff = $dtpDeadline.Value - (Get-Date)
            $deadlineOffset = [Math]::Max(0, [Math]::Round($diff.TotalHours))
        }

        $tmplPath = Join-Path (Join-Path $PSScriptRoot "Templates") "$($tmplName -replace '[^\w\s-]','').json"

        $timeBasisMap = @('LocalTime', 'Utc')
        Save-DeploymentTemplate -TemplatePath $tmplPath -TemplateName $tmplName -Config @{
            DeployPurpose               = $purpose
            UserNotification            = $notifMap[$cboNotification.SelectedIndex]
            TimeBasedOn                 = $timeBasisMap[$cboTimeBasis.SelectedIndex]
            OverrideServiceWindow       = $chkOverrideMW.Checked
            RebootOutsideServiceWindow  = $chkRebootOutside.Checked
            AllowMeteredConnection      = $chkMetered.Checked
            AllowBoundaryFallback       = $chkBoundaryFallback.Checked
            AllowMicrosoftUpdate        = $chkMicrosoftUpdate.Checked
            RequirePostRebootFullScan   = $chkPostRebootScan.Checked
            DefaultDeadlineOffsetHours  = $deadlineOffset
        }

        # Reload templates into combobox
        $script:Templates = Get-DeploymentTemplates -TemplatePath (Join-Path $PSScriptRoot "Templates")
        $cboTemplate.Items.Clear()
        [void]$cboTemplate.Items.Add("(None)")
        foreach ($t in $script:Templates) { [void]$cboTemplate.Items.Add($t.Name) }
        $cboTemplate.SelectedIndex = 0

        Add-LogLine -TextBox $txtLog -Message "Template saved: $tmplName"
    }
    $inputDlg.Dispose()
})

# ---------------------------------------------------------------------------
# Event: Connect
# ---------------------------------------------------------------------------

$btnConnect.Add_Click({
    if (-not $script:Prefs.SiteCode -or -not $script:Prefs.SMSProvider) {
        [System.Windows.Forms.MessageBox]::Show(
            "Configure Site Code and SMS Provider in File > Preferences first.",
            "Connection Required", "OK", "Warning") | Out-Null
        return
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblConnStatus.Text = "Connecting..."
    $lblConnStatus.ForeColor = $clrWarnText
    [System.Windows.Forms.Application]::DoEvents()

    $ok = Connect-CMSite -SiteCode $script:Prefs.SiteCode -SMSProvider $script:Prefs.SMSProvider

    if ($ok) {
        $lblConnStatus.Text = "Connected"
        $lblConnStatus.ForeColor = $clrOkText
        $statusLabel.Text = "Connected to site $($script:Prefs.SiteCode)"
        Add-LogLine -TextBox $txtLog -Message "Connected to site $($script:Prefs.SiteCode) on $($script:Prefs.SMSProvider)"

        # Populate DP groups
        $clbDPGroups.Items.Clear()
        $dpGroups = Get-DPGroupList
        foreach ($g in $dpGroups) { [void]$clbDPGroups.Items.Add($g.Name, $false) }
        $clbDPGroups.Enabled = ($dpGroups.Count -gt 0)
        Add-LogLine -TextBox $txtLog -Message "$($dpGroups.Count) DP group(s) loaded"
    } else {
        $lblConnStatus.Text = "Failed"
        $lblConnStatus.ForeColor = $clrErrText
        Add-LogLine -TextBox $txtLog -Message "Connection FAILED"
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# ---------------------------------------------------------------------------
# Event: Browse Application
# ---------------------------------------------------------------------------

$btnBrowseApp.Add_Click({
    if (-not (Test-CMConnection)) {
        [System.Windows.Forms.MessageBox]::Show("Connect to a site first.", "Not Connected", "OK", "Warning") | Out-Null
        return
    }

    $isSUG = ($cboDeployType.SelectedIndex -eq 1)
    if ($isSUG) {
        $selected = Show-SearchDialog -Title "Search Software Update Groups" -SearchLabel "SUG:" `
            -ColumnNames @('LocalizedDisplayName', 'NumberOfUpdates') `
            -NameColumn 'LocalizedDisplayName' `
            -SearchAction { param($term) Get-CMSoftwareUpdateGroup -Name "*$term*" -ErrorAction SilentlyContinue | Select-Object LocalizedDisplayName, NumberOfUpdates | Sort-Object LocalizedDisplayName }
    } else {
        $selected = Show-SearchDialog -Title "Search Applications" -SearchLabel "App:" `
            -ColumnNames @('LocalizedDisplayName', 'SoftwareVersion', 'PackageID') `
            -NameColumn 'LocalizedDisplayName' `
            -SearchAction { param($term) Search-CMApplicationByName -SearchText $term }
    }

    if ($selected) { $txtAppName.Text = $selected }
})

# ---------------------------------------------------------------------------
# Event: Browse Collection
# ---------------------------------------------------------------------------

$btnBrowseColl.Add_Click({
    if (-not (Test-CMConnection)) {
        [System.Windows.Forms.MessageBox]::Show("Connect to a site first.", "Not Connected", "OK", "Warning") | Out-Null
        return
    }

    $selected = Show-SearchDialog -Title "Search Device Collections" -SearchLabel "Collection:" `
        -ColumnNames @('Name', 'CollectionID', 'MemberCount') `
        -NameColumn 'Name' `
        -SearchAction { param($term) Search-CMCollectionByName -SearchText $term }

    if ($selected) { $txtCollName.Text = $selected }
})

# ---------------------------------------------------------------------------
# Event: Validate
# ---------------------------------------------------------------------------

$script:ValidatedApp = $null
$script:ValidatedCol = $null

$btnValidate.Add_Click({
    # Pre-checks
    if (-not (Test-CMConnection)) {
        [System.Windows.Forms.MessageBox]::Show("Connect to MECM first.", "Not Connected", "OK", "Warning") | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($txtAppName.Text)) {
        $typeLabel = if ($cboDeployType.SelectedIndex -eq 1) { 'SUG name' } else { 'application name' }
        [System.Windows.Forms.MessageBox]::Show("Enter an $typeLabel.", "Missing Input", "OK", "Warning") | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($txtCollName.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Enter a collection name.", "Missing Input", "OK", "Warning") | Out-Null
        return
    }

    $isSUG = ($cboDeployType.SelectedIndex -eq 1)
    $rtbValidation.Clear()
    $btnDeploy.Enabled = $false
    $script:ValidatedApp = $null
    $script:ValidatedCol = $null
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Add-LogLine -TextBox $txtLog -Message ("Validating {0} deployment..." -f $(if ($isSUG) { 'SUG' } else { 'application' }))
    [System.Windows.Forms.Application]::DoEvents()

    $allPassed = $true
    $targetObj = $null

    if ($isSUG) {
        # Check 1: SUG exists
        $targetObj = Test-SUGExists -SUGName $txtAppName.Text.Trim()
        if ($null -ne $targetObj) {
            Add-ValidationLine -Icon "[PASS]" -Message ("SUG found: {0} ({1} updates)" -f $targetObj.LocalizedDisplayName, $targetObj.NumberOfUpdates) -Color $clrOkText
            if ($targetObj.NumberOfUpdates -eq 0) {
                Add-ValidationLine -Icon "[WARN]" -Message "SUG contains 0 updates" -Color $clrErrText
            }
        } else {
            Add-ValidationLine -Icon "[FAIL]" -Message ("Software Update Group not found: {0}" -f $txtAppName.Text.Trim()) -Color $clrErrText
            $allPassed = $false
        }
        [System.Windows.Forms.Application]::DoEvents()

        # Check 2: Skipped for SUG (no content distribution check)
        Add-ValidationLine -Icon "[SKIP]" -Message "Content distribution check not applicable for SUGs" -Color $clrText
        [System.Windows.Forms.Application]::DoEvents()
    } else {
        # Check 1: Application exists
        $targetObj = Test-ApplicationExists -ApplicationName $txtAppName.Text.Trim()
        if ($null -ne $targetObj) {
            Add-ValidationLine -Icon "[PASS]" -Message ("Application found: {0} v{1}" -f $targetObj.LocalizedDisplayName, $targetObj.SoftwareVersion) -Color $clrOkText
        } else {
            Add-ValidationLine -Icon "[FAIL]" -Message ("Application not found: {0}" -f $txtAppName.Text.Trim()) -Color $clrErrText
            $allPassed = $false
        }
        [System.Windows.Forms.Application]::DoEvents()

        # Check 2: Content distributed (only if app found)
        if ($null -ne $targetObj) {
            $distStatus = Test-ContentDistributed -Application $targetObj
            if ($distStatus.IsFullyDistributed) {
                Add-ValidationLine -Icon "[PASS]" -Message ("Content distributed: {0}/{1} DPs" -f $distStatus.NumberSuccess, $distStatus.Targeted) -Color $clrOkText
            } else {
                Add-ValidationLine -Icon "[FAIL]" -Message ("Content NOT fully distributed: {0}/{1} success, {2} errors" -f $distStatus.NumberSuccess, $distStatus.Targeted, $distStatus.NumberErrors) -Color $clrErrText
                $allPassed = $false
            }
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    # Check 3: Collection valid
    $col = Test-CollectionValid -CollectionName $txtCollName.Text.Trim()
    if ($null -ne $col) {
        Add-ValidationLine -Icon "[PASS]" -Message ("Collection found: {0} (ID: {1}, {2} members)" -f $col.Name, $col.CollectionID, $col.MemberCount) -Color $clrOkText
    } else {
        Add-ValidationLine -Icon "[FAIL]" -Message ("Collection not found or not a Device collection: {0}" -f $txtCollName.Text.Trim()) -Color $clrErrText
        $allPassed = $false
    }
    [System.Windows.Forms.Application]::DoEvents()

    # Check 4: Collection safe (only if found)
    if ($null -ne $col) {
        $safety = Test-CollectionSafe -Collection $col
        if ($safety.IsSafe) {
            Add-ValidationLine -Icon "[PASS]" -Message "Collection passed safety check" -Color $clrOkText
        } else {
            Add-ValidationLine -Icon "[FAIL]" -Message ("BLOCKED: {0}" -f $safety.Reason) -Color $clrErrText
            $allPassed = $false
        }
    }
    [System.Windows.Forms.Application]::DoEvents()

    # Check 5: Duplicate deployment (skip for SUG)
    if (-not $isSUG -and $null -ne $targetObj -and $null -ne $col) {
        $dupe = Test-DuplicateDeployment -ApplicationName $txtAppName.Text.Trim() -CollectionName $txtCollName.Text.Trim()
        if ($null -eq $dupe) {
            Add-ValidationLine -Icon "[PASS]" -Message "No duplicate deployment exists" -Color $clrOkText
        } else {
            Add-ValidationLine -Icon "[FAIL]" -Message "Duplicate deployment already exists for this app/collection" -Color $clrErrText
            $allPassed = $false
        }
    } elseif ($isSUG) {
        Add-ValidationLine -Icon "[SKIP]" -Message "Duplicate check not applicable for SUGs" -Color $clrText
    }
    [System.Windows.Forms.Application]::DoEvents()

    # Summary
    if ($allPassed -and $null -ne $targetObj -and $null -ne $col) {
        $deployType = if ($isSUG) { 'SUG' } else { 'Application' }
        $preview = Get-DeploymentPreview -TargetObject $targetObj -Collection $col -DeploymentType $deployType
        Add-ValidationLine -Icon "" -Message "" -Color $clrText
        Add-ValidationLine -Icon "[INFO]" -Message ("Impact: {0} {1} -> {2} ({3} devices)" -f $preview.ApplicationName, $preview.ApplicationVersion, $preview.CollectionName, $preview.MemberCount) -Color $clrAccent
        $btnDeploy.Enabled = $true
        $script:ValidatedApp = $targetObj
        $script:ValidatedCol = $col
        Add-LogLine -TextBox $txtLog -Message "Validation PASSED - Deploy button enabled"
    } else {
        Add-ValidationLine -Icon "" -Message "" -Color $clrText
        Add-ValidationLine -Icon "[INFO]" -Message "Validation FAILED - fix errors above before deploying" -Color $clrErrText
        Add-LogLine -TextBox $txtLog -Message "Validation FAILED"
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# ---------------------------------------------------------------------------
# Event: Deploy
# ---------------------------------------------------------------------------

$btnDeploy.Add_Click({
    if ($null -eq $script:ValidatedApp -or $null -eq $script:ValidatedCol) { return }

    $isSUG = ($cboDeployType.SelectedIndex -eq 1)
    $purpose = if ($radRequired.Checked) { 'Required' } else { 'Available' }
    $deployType = if ($isSUG) { 'SUG' } else { 'Application' }
    $preview = Get-DeploymentPreview -TargetObject $script:ValidatedApp -Collection $script:ValidatedCol -DeploymentType $deployType

    # Map notification and time basis comboboxes to parameter values
    $notifMap = @('DisplayAll', 'DisplaySoftwareCenterOnly', 'HideAll')
    $notifValue = $notifMap[$cboNotification.SelectedIndex]
    $timeBasisMap = @('LocalTime', 'Utc')
    $timeBasisValue = $timeBasisMap[$cboTimeBasis.SelectedIndex]

    # Confirmation dialog
    $timeStr = if ($cboTimeBasis.SelectedIndex -eq 1) { " (UTC)" } else { "" }
    $deadlineStr = if ($radRequired.Checked) { "`nDeadline: $($dtpDeadline.Value.ToString('yyyy-MM-dd HH:mm'))$timeStr" } else { '' }
    $meteredStr = if ($chkMetered.Checked) { "`nMetered connection: Allowed" } else { '' }
    $confirmMsg = ("Deploy {0} {1} to {2} ({3} devices) as {4}?{5}{6}" -f
        $preview.ApplicationName, $preview.ApplicationVersion,
        $preview.CollectionName, $preview.MemberCount,
        $purpose, $deadlineStr, $meteredStr)

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        $confirmMsg, "Confirm Deployment",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    # Distribute content to selected DP groups (if any checked)
    $selectedDPGroups = @()
    for ($i = 0; $i -lt $clbDPGroups.Items.Count; $i++) {
        if ($clbDPGroups.GetItemChecked($i)) { $selectedDPGroups += $clbDPGroups.Items[$i] }
    }
    if ($selectedDPGroups.Count -gt 0 -and -not $isSUG) {
        Add-LogLine -TextBox $txtLog -Message ("Distributing content to {0} DP group(s)..." -f $selectedDPGroups.Count)
        [System.Windows.Forms.Application]::DoEvents()
        $distResults = Start-ContentDistributionToGroups -Application $script:ValidatedApp -DPGroupNames $selectedDPGroups
        foreach ($r in $distResults) {
            if ($r.AlreadyTargeted) {
                Add-LogLine -TextBox $txtLog -Message "  $($r.Group): already targeted"
            } elseif ($r.Success) {
                Add-LogLine -TextBox $txtLog -Message "  $($r.Group): distribution started"
            } else {
                Add-LogLine -TextBox $txtLog -Message "  $($r.Group): FAILED - $($r.Error)"
            }
        }
    }

    Add-LogLine -TextBox $txtLog -Message ("Executing {0} deployment..." -f $deployType)
    [System.Windows.Forms.Application]::DoEvents()

    if ($isSUG) {
        $deployParams = @{
            SUG                        = $script:ValidatedApp
            Collection                 = $script:ValidatedCol
            DeployPurpose              = $purpose
            AvailableDateTime          = $dtpAvailable.Value
            TimeBasedOn                = $timeBasisValue
            UserNotification           = $notifValue
            SoftwareInstallation       = $chkOverrideMW.Checked
            AllowRestart               = $chkRebootOutside.Checked
            UseMeteredNetwork          = $chkMetered.Checked
            AllowBoundaryFallback      = $chkBoundaryFallback.Checked
            DownloadFromMicrosoftUpdate = $chkMicrosoftUpdate.Checked
            RequirePostRebootFullScan  = $chkPostRebootScan.Checked
        }
        if ($radRequired.Checked) {
            $deployParams['DeadlineDateTime'] = $dtpDeadline.Value
        }
        $result = Invoke-SUGDeployment @deployParams
    } else {
        $deployParams = @{
            Application                 = $script:ValidatedApp
            Collection                  = $script:ValidatedCol
            DeployPurpose               = $purpose
            AvailableDateTime           = $dtpAvailable.Value
            TimeBasedOn                 = $timeBasisValue
            UserNotification            = $notifValue
            OverrideServiceWindow       = $chkOverrideMW.Checked
            RebootOutsideServiceWindow  = $chkRebootOutside.Checked
            UseMeteredNetwork           = $chkMetered.Checked
        }
        if ($radRequired.Checked) {
            $deployParams['DeadlineDateTime'] = $dtpDeadline.Value
        }
        $result = Invoke-ApplicationDeployment @deployParams
    }

    # Resolve deployment log path
    $logPath = if ($script:Prefs.DeploymentLogPath) {
        Join-Path $script:Prefs.DeploymentLogPath "deployment-log.jsonl"
    } else {
        Join-Path $PSScriptRoot "Logs\deployment-log.jsonl"
    }

    if ($result.Success) {
        Add-ValidationLine -Icon "" -Message "" -Color $clrText
        Add-ValidationLine -Icon "[OK]" -Message ("Deployment SUCCEEDED (ID: {0})" -f $result.DeploymentID) -Color $clrOkText
        Add-LogLine -TextBox $txtLog -Message ("Deployment succeeded: ID {0}" -f $result.DeploymentID)
        $statusLabel.Text = ("Last deployment: {0} -> {1} (ID: {2})" -f $preview.ApplicationName, $preview.CollectionName, $result.DeploymentID)

        Write-DeploymentLog -LogPath $logPath -Record @{
            DeploymentType     = $deployType
            ApplicationName    = $preview.ApplicationName
            ApplicationVersion = $preview.ApplicationVersion
            CollectionName     = $preview.CollectionName
            CollectionID       = $preview.CollectionID
            MemberCount        = $preview.MemberCount
            DeployPurpose      = $purpose
            DeadlineDateTime   = if ($radRequired.Checked) { $dtpDeadline.Value.ToString('yyyy-MM-ddTHH:mm:ss') } else { '' }
            DeploymentID       = $result.DeploymentID
            Result             = 'Success'
        }
    } else {
        Add-ValidationLine -Icon "" -Message "" -Color $clrText
        Add-ValidationLine -Icon "[FAIL]" -Message ("Deployment FAILED: {0}" -f $result.Error) -Color $clrErrText
        Add-LogLine -TextBox $txtLog -Message ("Deployment FAILED: {0}" -f $result.Error)

        Write-DeploymentLog -LogPath $logPath -Record @{
            DeploymentType     = $deployType
            ApplicationName    = $txtAppName.Text.Trim()
            ApplicationVersion = ''
            CollectionName     = $txtCollName.Text.Trim()
            CollectionID       = ''
            MemberCount        = 0
            DeployPurpose      = $purpose
            DeadlineDateTime   = ''
            DeploymentID       = ''
            Result             = "Failed: $($result.Error)"
        }
    }

    $btnDeploy.Enabled = $false
    $script:ValidatedApp = $null
    $script:ValidatedCol = $null
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
})

# ---------------------------------------------------------------------------
# Event: Export CSV
# ---------------------------------------------------------------------------

$btnExportCsv.Add_Click({
    $logPath = if ($script:Prefs.DeploymentLogPath) {
        Join-Path $script:Prefs.DeploymentLogPath "deployment-log.jsonl"
    } else {
        Join-Path $PSScriptRoot "Logs\deployment-log.jsonl"
    }

    $records = Get-DeploymentHistory -LogPath $logPath
    if ($records.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No deployment history found.", "No Data", "OK", "Information") | Out-Null
        return
    }

    $reportsDir = Join-Path $PSScriptRoot "Reports"
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "CSV Files (*.csv)|*.csv"
    $sfd.FileName = "DeploymentHistory-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-DeploymentHistoryCsv -Records $records -OutputPath $sfd.FileName
        Add-LogLine -TextBox $txtLog -Message "Exported CSV: $($sfd.FileName)"
    }
})

# ---------------------------------------------------------------------------
# Event: Export HTML
# ---------------------------------------------------------------------------

$btnExportHtml.Add_Click({
    $logPath = if ($script:Prefs.DeploymentLogPath) {
        Join-Path $script:Prefs.DeploymentLogPath "deployment-log.jsonl"
    } else {
        Join-Path $PSScriptRoot "Logs\deployment-log.jsonl"
    }

    $records = Get-DeploymentHistory -LogPath $logPath
    if ($records.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No deployment history found.", "No Data", "OK", "Information") | Out-Null
        return
    }

    $reportsDir = Join-Path $PSScriptRoot "Reports"
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "HTML Files (*.html)|*.html"
    $sfd.FileName = "DeploymentHistory-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-DeploymentHistoryHtml -Records $records -OutputPath $sfd.FileName
        Add-LogLine -TextBox $txtLog -Message "Exported HTML: $($sfd.FileName)"
    }
})

# ---------------------------------------------------------------------------
# Window state + run
# ---------------------------------------------------------------------------

$form.Add_Shown({ Restore-WindowState })
$form.Add_FormClosing({
    Save-WindowState
    Disconnect-CMSite
})

Add-LogLine -TextBox $txtLog -Message "Deployment Helper started. Configure site in File > Preferences, then click Connect."

[void]$form.ShowDialog()
$form.Dispose()
