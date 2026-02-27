<#
.SYNOPSIS
    WinForms front-end for Deployment Helper - safe MECM application deployment.

.DESCRIPTION
    Single-pane deployment tool with pre-execution validation, safety guardrails,
    and immutable deployment audit logging.

    Features:
      - Enter change ticket, application, collection
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
            $form.Size = New-Object System.Drawing.Size($state.Width, $state.Height)
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
                Start-Process powershell -ArgumentList @('-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
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
    $lblVersion.Text = "Deployment Helper v1.0.0"
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
$form.Size = New-Object System.Drawing.Size(900, 820)
$form.MinimumSize = New-Object System.Drawing.Size(780, 720)
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
$pnlForm.Height = 340
$pnlForm.BackColor = $clrPanelBg
$form.Controls.Add($pnlForm)

# Load templates
$script:Templates = Get-DeploymentTemplates -TemplatePath (Join-Path $PSScriptRoot "Templates")

# Row 1: Change Ticket #
$lblTicket = New-Object System.Windows.Forms.Label
$lblTicket.Text = "Change Ticket #:"
$lblTicket.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblTicket.ForeColor = $clrText
$lblTicket.Location = New-Object System.Drawing.Point(14, 14)
$lblTicket.AutoSize = $true
$pnlForm.Controls.Add($lblTicket)

$txtTicket = New-Object System.Windows.Forms.TextBox
$txtTicket.SetBounds(160, 11, 200, 24)
$txtTicket.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtTicket.BackColor = $clrDetailBg
$txtTicket.ForeColor = $clrText
$txtTicket.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
$pnlForm.Controls.Add($txtTicket)

# Row 2: Application
$lblAppName = New-Object System.Windows.Forms.Label
$lblAppName.Text = "Application:"
$lblAppName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblAppName.ForeColor = $clrText
$lblAppName.Location = New-Object System.Drawing.Point(14, 46)
$lblAppName.AutoSize = $true
$pnlForm.Controls.Add($lblAppName)

$txtAppName = New-Object System.Windows.Forms.TextBox
$txtAppName.SetBounds(160, 43, 300, 24)
$txtAppName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtAppName.BackColor = $clrDetailBg
$txtAppName.ForeColor = $clrText
$txtAppName.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
$pnlForm.Controls.Add($txtAppName)

# Row 3: Collection
$lblCollName = New-Object System.Windows.Forms.Label
$lblCollName.Text = "Collection:"
$lblCollName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblCollName.ForeColor = $clrText
$lblCollName.Location = New-Object System.Drawing.Point(14, 78)
$lblCollName.AutoSize = $true
$pnlForm.Controls.Add($lblCollName)

$txtCollName = New-Object System.Windows.Forms.TextBox
$txtCollName.SetBounds(160, 75, 300, 24)
$txtCollName.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtCollName.BackColor = $clrDetailBg
$txtCollName.ForeColor = $clrText
$txtCollName.BorderStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.BorderStyle]::None } else { [System.Windows.Forms.BorderStyle]::FixedSingle }
$pnlForm.Controls.Add($txtCollName)

# Row 4: Template
$lblTemplate = New-Object System.Windows.Forms.Label
$lblTemplate.Text = "Template:"
$lblTemplate.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblTemplate.ForeColor = $clrText
$lblTemplate.Location = New-Object System.Drawing.Point(14, 110)
$lblTemplate.AutoSize = $true
$pnlForm.Controls.Add($lblTemplate)

$cboTemplate = New-Object System.Windows.Forms.ComboBox
$cboTemplate.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboTemplate.SetBounds(160, 107, 220, 24)
$cboTemplate.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboTemplate.BackColor = $clrDetailBg
$cboTemplate.ForeColor = $clrText
$cboTemplate.FlatStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.FlatStyle]::Flat } else { [System.Windows.Forms.FlatStyle]::Standard }
[void]$cboTemplate.Items.Add("(None)")
foreach ($tmpl in $script:Templates) { [void]$cboTemplate.Items.Add($tmpl.Name) }
$cboTemplate.SelectedIndex = 0
$pnlForm.Controls.Add($cboTemplate)

# Row 5: Purpose
$lblPurpose = New-Object System.Windows.Forms.Label
$lblPurpose.Text = "Purpose:"
$lblPurpose.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblPurpose.ForeColor = $clrText
$lblPurpose.Location = New-Object System.Drawing.Point(14, 142)
$lblPurpose.AutoSize = $true
$pnlForm.Controls.Add($lblPurpose)

$radAvailable = New-Object System.Windows.Forms.RadioButton
$radAvailable.Text = "Available"
$radAvailable.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radAvailable.ForeColor = $clrText
$radAvailable.BackColor = $clrPanelBg
$radAvailable.Location = New-Object System.Drawing.Point(160, 140)
$radAvailable.AutoSize = $true
$radAvailable.Checked = $true
if ($script:Prefs.DarkMode) { $radAvailable.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $radAvailable.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($radAvailable)

$radRequired = New-Object System.Windows.Forms.RadioButton
$radRequired.Text = "Required"
$radRequired.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$radRequired.ForeColor = $clrText
$radRequired.BackColor = $clrPanelBg
$radRequired.Location = New-Object System.Drawing.Point(270, 140)
$radRequired.AutoSize = $true
if ($script:Prefs.DarkMode) { $radRequired.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $radRequired.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($radRequired)

# Row 6: Available date
$lblAvailable = New-Object System.Windows.Forms.Label
$lblAvailable.Text = "Available:"
$lblAvailable.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblAvailable.ForeColor = $clrText
$lblAvailable.Location = New-Object System.Drawing.Point(14, 174)
$lblAvailable.AutoSize = $true
$pnlForm.Controls.Add($lblAvailable)

$dtpAvailable = New-Object System.Windows.Forms.DateTimePicker
$dtpAvailable.SetBounds(160, 171, 200, 24)
$dtpAvailable.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$dtpAvailable.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpAvailable.CustomFormat = "yyyy-MM-dd HH:mm"
$dtpAvailable.Value = Get-Date
$pnlForm.Controls.Add($dtpAvailable)

# Row 7: Deadline date
$lblDeadline = New-Object System.Windows.Forms.Label
$lblDeadline.Text = "Deadline:"
$lblDeadline.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblDeadline.ForeColor = $clrText
$lblDeadline.Location = New-Object System.Drawing.Point(14, 206)
$lblDeadline.AutoSize = $true
$pnlForm.Controls.Add($lblDeadline)

$dtpDeadline = New-Object System.Windows.Forms.DateTimePicker
$dtpDeadline.SetBounds(160, 203, 200, 24)
$dtpDeadline.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$dtpDeadline.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$dtpDeadline.CustomFormat = "yyyy-MM-dd HH:mm"
$dtpDeadline.Value = (Get-Date).AddHours(24)
$dtpDeadline.Enabled = $false
$pnlForm.Controls.Add($dtpDeadline)

# Row 8: Notification
$lblNotification = New-Object System.Windows.Forms.Label
$lblNotification.Text = "Notification:"
$lblNotification.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblNotification.ForeColor = $clrText
$lblNotification.Location = New-Object System.Drawing.Point(14, 238)
$lblNotification.AutoSize = $true
$pnlForm.Controls.Add($lblNotification)

$cboNotification = New-Object System.Windows.Forms.ComboBox
$cboNotification.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboNotification.SetBounds(160, 235, 260, 24)
$cboNotification.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cboNotification.BackColor = $clrDetailBg
$cboNotification.ForeColor = $clrText
$cboNotification.FlatStyle = if ($script:Prefs.DarkMode) { [System.Windows.Forms.FlatStyle]::Flat } else { [System.Windows.Forms.FlatStyle]::Standard }
[void]$cboNotification.Items.AddRange(@('Display All Notifications', 'Display in Software Center Only', 'Hide All Notifications'))
$cboNotification.SelectedIndex = 0
$pnlForm.Controls.Add($cboNotification)

# Row 9: Maintenance window overrides
$chkOverrideMW = New-Object System.Windows.Forms.CheckBox
$chkOverrideMW.Text = "Allow outside maintenance window"
$chkOverrideMW.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkOverrideMW.ForeColor = $clrText
$chkOverrideMW.BackColor = $clrPanelBg
$chkOverrideMW.Location = New-Object System.Drawing.Point(160, 268)
$chkOverrideMW.AutoSize = $true
if ($script:Prefs.DarkMode) { $chkOverrideMW.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkOverrideMW.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($chkOverrideMW)

$chkRebootOutside = New-Object System.Windows.Forms.CheckBox
$chkRebootOutside.Text = "Reboot outside maintenance window"
$chkRebootOutside.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkRebootOutside.ForeColor = $clrText
$chkRebootOutside.BackColor = $clrPanelBg
$chkRebootOutside.Location = New-Object System.Drawing.Point(160, 292)
$chkRebootOutside.AutoSize = $true
if ($script:Prefs.DarkMode) { $chkRebootOutside.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkRebootOutside.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
$pnlForm.Controls.Add($chkRebootOutside)

# Row 10: Validate + Deploy buttons
$btnValidate = New-Object System.Windows.Forms.Button
$btnValidate.Text = "Validate"
$btnValidate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnValidate.Size = New-Object System.Drawing.Size(120, 32)
$btnValidate.Location = New-Object System.Drawing.Point(160, 302)
Set-ModernButtonStyle -Button $btnValidate -BackColor $clrAccent
$pnlForm.Controls.Add($btnValidate)

$btnDeploy = New-Object System.Windows.Forms.Button
$btnDeploy.Text = "Deploy"
$btnDeploy.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnDeploy.Size = New-Object System.Drawing.Size(120, 32)
$btnDeploy.Location = New-Object System.Drawing.Point(290, 302)
$btnDeploy.Enabled = $false
Set-ModernButtonStyle -Button $btnDeploy -BackColor ([System.Drawing.Color]::FromArgb(34, 139, 34))
$pnlForm.Controls.Add($btnDeploy)

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

# BringToFront processes innermost-last. Top-docked in reverse visual order:
$pnlSep2.BringToFront()
$pnlForm.BringToFront()
$pnlSep1.BringToFront()
$pnlConnBar.BringToFront()
$pnlHeader.BringToFront()

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
    if ($tmpl.DefaultDeadlineOffsetHours -and $tmpl.DefaultDeadlineOffsetHours -gt 0) {
        $dtpDeadline.Value = (Get-Date).AddHours($tmpl.DefaultDeadlineOffsetHours)
    }
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
    } else {
        $lblConnStatus.Text = "Failed"
        $lblConnStatus.ForeColor = $clrErrText
        Add-LogLine -TextBox $txtLog -Message "Connection FAILED"
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::Default
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
        [System.Windows.Forms.MessageBox]::Show("Enter an application name.", "Missing Input", "OK", "Warning") | Out-Null
        return
    }
    if ([string]::IsNullOrWhiteSpace($txtCollName.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Enter a collection name.", "Missing Input", "OK", "Warning") | Out-Null
        return
    }

    $rtbValidation.Clear()
    $btnDeploy.Enabled = $false
    $script:ValidatedApp = $null
    $script:ValidatedCol = $null
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Add-LogLine -TextBox $txtLog -Message "Validating deployment..."
    [System.Windows.Forms.Application]::DoEvents()

    $allPassed = $true

    # Check 1: Application exists
    $app = Test-ApplicationExists -ApplicationName $txtAppName.Text.Trim()
    if ($null -ne $app) {
        Add-ValidationLine -Icon "[PASS]" -Message ("Application found: {0} v{1}" -f $app.LocalizedDisplayName, $app.SoftwareVersion) -Color $clrOkText
    } else {
        Add-ValidationLine -Icon "[FAIL]" -Message ("Application not found: {0}" -f $txtAppName.Text.Trim()) -Color $clrErrText
        $allPassed = $false
    }
    [System.Windows.Forms.Application]::DoEvents()

    # Check 2: Content distributed (only if app found)
    if ($null -ne $app) {
        $distStatus = Test-ContentDistributed -Application $app
        if ($distStatus.IsFullyDistributed) {
            Add-ValidationLine -Icon "[PASS]" -Message ("Content distributed: {0}/{1} DPs" -f $distStatus.NumberSuccess, $distStatus.Targeted) -Color $clrOkText
        } else {
            Add-ValidationLine -Icon "[FAIL]" -Message ("Content NOT fully distributed: {0}/{1} success, {2} errors" -f $distStatus.NumberSuccess, $distStatus.Targeted, $distStatus.NumberErrors) -Color $clrErrText
            $allPassed = $false
        }
    }
    [System.Windows.Forms.Application]::DoEvents()

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

    # Check 5: Duplicate deployment
    if ($null -ne $app -and $null -ne $col) {
        $dupe = Test-DuplicateDeployment -ApplicationName $txtAppName.Text.Trim() -CollectionName $txtCollName.Text.Trim()
        if ($null -eq $dupe) {
            Add-ValidationLine -Icon "[PASS]" -Message "No duplicate deployment exists" -Color $clrOkText
        } else {
            Add-ValidationLine -Icon "[FAIL]" -Message "Duplicate deployment already exists for this app/collection" -Color $clrErrText
            $allPassed = $false
        }
    }
    [System.Windows.Forms.Application]::DoEvents()

    # Summary
    if ($allPassed -and $null -ne $app -and $null -ne $col) {
        $preview = Get-DeploymentPreview -Application $app -Collection $col
        Add-ValidationLine -Icon "" -Message "" -Color $clrText
        Add-ValidationLine -Icon "[INFO]" -Message ("Impact: {0} v{1} -> {2} ({3} devices)" -f $preview.ApplicationName, $preview.ApplicationVersion, $preview.CollectionName, $preview.MemberCount) -Color $clrAccent
        $btnDeploy.Enabled = $true
        $script:ValidatedApp = $app
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

    $purpose = if ($radRequired.Checked) { 'Required' } else { 'Available' }
    $preview = Get-DeploymentPreview -Application $script:ValidatedApp -Collection $script:ValidatedCol

    # Map notification combobox to parameter value
    $notifMap = @('DisplayAll', 'DisplaySoftwareCenterOnly', 'HideAll')
    $notifValue = $notifMap[$cboNotification.SelectedIndex]

    # Confirmation dialog
    $deadlineStr = if ($radRequired.Checked) { "`nDeadline: $($dtpDeadline.Value.ToString('yyyy-MM-dd HH:mm'))" } else { '' }
    $confirmMsg = ("Deploy {0} v{1} to {2} ({3} devices) as {4}?{5}" -f
        $preview.ApplicationName, $preview.ApplicationVersion,
        $preview.CollectionName, $preview.MemberCount,
        $purpose, $deadlineStr)

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        $confirmMsg, "Confirm Deployment",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)

    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Add-LogLine -TextBox $txtLog -Message "Executing deployment..."
    [System.Windows.Forms.Application]::DoEvents()

    $deployParams = @{
        Application                 = $script:ValidatedApp
        Collection                  = $script:ValidatedCol
        DeployPurpose               = $purpose
        AvailableDateTime           = $dtpAvailable.Value
        UserNotification            = $notifValue
        OverrideServiceWindow       = $chkOverrideMW.Checked
        RebootOutsideServiceWindow  = $chkRebootOutside.Checked
        Comment                     = $txtTicket.Text.Trim()
    }
    if ($radRequired.Checked) {
        $deployParams['DeadlineDateTime'] = $dtpDeadline.Value
    }

    $result = Invoke-ApplicationDeployment @deployParams

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
            ChangeTicket       = $txtTicket.Text.Trim()
            ApplicationName    = $preview.ApplicationName
            ApplicationVersion = $preview.ApplicationVersion
            CollectionName     = $preview.CollectionName
            CollectionID       = $preview.CollectionID
            MemberCount        = $preview.MemberCount
            DeployPurpose      = $purpose
            DeadlineDateTime   = if ($radRequired.Checked) { $dtpDeadline.Value.ToString('yyyy-MM-ddTHH:mm:ss') } else { '' }
            DeploymentID       = $result.DeploymentID
            Result             = 'Success'
            Comment            = $txtTicket.Text.Trim()
        }
    } else {
        Add-ValidationLine -Icon "" -Message "" -Color $clrText
        Add-ValidationLine -Icon "[FAIL]" -Message ("Deployment FAILED: {0}" -f $result.Error) -Color $clrErrText
        Add-LogLine -TextBox $txtLog -Message ("Deployment FAILED: {0}" -f $result.Error)

        Write-DeploymentLog -LogPath $logPath -Record @{
            ChangeTicket       = $txtTicket.Text.Trim()
            ApplicationName    = $txtAppName.Text.Trim()
            ApplicationVersion = ''
            CollectionName     = $txtCollName.Text.Trim()
            CollectionID       = ''
            MemberCount        = 0
            DeployPurpose      = $purpose
            DeadlineDateTime   = ''
            DeploymentID       = ''
            Result             = "Failed: $($result.Error)"
            Comment            = $txtTicket.Text.Trim()
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
