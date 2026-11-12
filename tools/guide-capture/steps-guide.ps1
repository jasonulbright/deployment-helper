$win = Start-App 'start-deploymenthelper.ps1' 'Deployment Helper'
Size-Win $win 1440 900
$win = Find-Top 'Deployment Helper' 10
function Rows($grid) { @((Get-El $win $grid).FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem)))) }
function WinNames { $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window); (@($root.FindAll('Descendants', $c)) | ForEach-Object { $_.Current.Name }) -join ' ; ' }
# Open a dialog with a button, capture the dialog, dump its tree and close it with its Cancel button.
# The step never answers OK or Yes. When the dialog has no Cancel button, the step fails.
function Dlg([string]$btn, [string]$title, [string]$shot, [int]$wait = 3) {
  Invoke-El (Get-El $win $btn); Start-Sleep -Seconds $wait
  $d = Find-Last $title 20
  Shot $d $shot; Dump-Tree $d $shot
  $cancel = Get-El $d 'btnCancel' -Quiet
  if (-not $cancel) { throw ('No Cancel button in ' + $title) }
  Invoke-El $cancel; Start-Sleep -Seconds 1
}
function Pick-Nav($dialog, [string]$name) { Select-El (Find-Role $dialog 'ListItem' $name); Start-Sleep -Milliseconds 700 }
Section 'first run' {
  Shot $win 'start-first-run'
  Invoke-El (Get-El $win 'btnOptions'); $d = Find-Last 'Options' 10; Shot $d 'options-empty'
  Set-Text (Get-El $d 'txtSiteCode') $Capture.SiteCode; Set-Text (Get-El $d 'txtSmsProvider') $Capture.SMSProvider; Shot $d 'options-connection'
  foreach ($p in @(@('Logging', 'options-logging'), @('History', 'options-history'), @('Templates', 'options-templates'), @('About', 'options-about'))) { Pick-Nav $d $p[0]; Shot $d $p[1] }
  Dump-Tree $d 'options-pages'
  Pick-Nav $d 'Connection'; Invoke-El (Get-El $d 'btnOK'); Start-Sleep -Seconds 2
}
Section 'apps form' {
  Shot $win 'apps-empty'
  Expand-El (Get-El $win 'cboApplyTemplate'); Start-Sleep -Milliseconds 600; Shot $win 'apps-template-list' -Hover
  $li = Find-Role $win 'ListItem' 'Workstation Pilot'; if ($li) { Click-El $li; Start-Sleep -Seconds 1; Shot $win 'apps-template-applied' } else { Log 'ERROR: template Workstation Pilot not found'; Collapse-El (Get-El $win 'cboApplyTemplate') }
}
Section 'browse dialogs' {
  Dlg 'btnBrowseTarget' 'Browse Applications' 'dialog-browse-application' 8
  # The light theme of the same dialog, with a selected row: white text on the accent blue.
  Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
  Invoke-El (Get-El $win 'btnBrowseTarget'); Start-Sleep -Seconds 3
  $d = Find-Last 'Browse Applications' 20
  $rows = @($d.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem))))
  if ($rows.Count) { Select-El $rows[0]; Start-Sleep -Milliseconds 700 } else { Log 'ERROR: the browse dialog shows no row to select' }
  Shot $d 'light-dialog-browse-application'
  $cancel = Get-El $d 'btnCancel' -Quiet
  if (-not $cancel) { throw 'No Cancel button in Browse Applications' }
  Invoke-El $cancel; Start-Sleep -Seconds 1
  Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
  Dlg 'btnBrowseCollection' 'Pick Device Collection' 'dialog-browse-collection' 12
}
Section 'apps validate' {
  Set-Text (Get-El $win 'txtTargetName') 'Doc Demo Viewer 1.0'; Set-Text (Get-El $win 'txtCollection') 'All Workstations'
  Select-El (Get-El $win 'radAvailable'); Start-Sleep -Milliseconds 500; Shot $win 'apps-filled'
  Invoke-El (Get-El $win 'btnValidate'); Start-Sleep -Seconds 12; Shot $win 'apps-validated'
  Log ('deploy enabled: ' + (Get-El $win 'btnDeploy').Current.IsEnabled)
}
Section 'apps required' {
  Select-El (Get-El $win 'radRequired'); Start-Sleep -Seconds 1; Shot $win 'apps-required'
}
# Click Deploy and capture the dialog that shows. The step answers only with the safe button:
# Cancel on Confirm Deployment, No on Check Deployment Schedule, OK on Deployment Blocked.
# Any other dialog stops the step. The guide never confirms a deployment.
function Deploy-AndAnswer([string]$shot) {
  $btn = Get-El $win 'btnDeploy'
  # The Deploy button is hidden until all five checks pass. Never invoke a hidden button.
  if ($btn.Current.IsOffscreen -or -not $btn.Current.IsEnabled) { Log ('The Deploy button is not shown, so no dialog was captured for ' + $shot + '.'); return }
  Invoke-El $btn; Start-Sleep -Seconds 5
  $safe = @{ 'Confirm Deployment' = 'Cancel'; 'Check Deployment Schedule' = 'No'; 'Deployment Blocked' = 'OK' }
  $found = $null
  foreach ($title in $safe.Keys) { $d = @(Find-All $title) | Select-Object -First 1; if ($d) { $found = @{ Window = $d; Title = $title }; break } }
  if (-not $found) { Log ('ERROR: no deployment dialog showed. Windows: ' + (WinNames)); return }
  Shot $found.Window $shot; Dump-Tree $found.Window $shot
  $answer = Get-El $found.Window $safe[$found.Title] -Quiet
  if (-not $answer) { throw ('No ' + $safe[$found.Title] + ' button in ' + $found.Title + '. Nothing was answered.') }
  Invoke-El $answer; Start-Sleep -Seconds 2
  if (Find-All $found.Title) { throw ('The dialog ' + $found.Title + ' is still open after ' + $safe[$found.Title]) }
}
Section 'deploy dialog with passing checks' {
  Set-Text (Get-El $win 'txtTargetName') 'CMCO-CANARY-App'; Start-Sleep -Milliseconds 500
  Invoke-El (Get-El $win 'btnValidate'); Start-Sleep -Seconds 12; Shot $win 'apps-validated-pass'
  Deploy-AndAnswer 'dialog-confirm-deployment'
}
Section 'dp picker' {
  Dlg 'btnDPPicker' 'Select DP Groups' 'dialog-dp-groups' 8
}
Section 'other types' {
  foreach ($t in @(@('btnPackages', 'packages-form'), @('btnTaskSequences', 'task-sequences-form'), @('btnSUG', 'sug-form'))) { Invoke-El (Get-El $win $t[0]); Start-Sleep -Seconds 2; Shot $win $t[1] }
}
Section 'rings' {
  Invoke-El (Get-El $win 'btnRings'); Start-Sleep -Seconds 3; Shot $win 'rings-empty'
  Dump-Tree $win 'rings'
  Expand-El (Get-El $win 'cboRingType'); Start-Sleep -Milliseconds 600; Shot $win 'rings-type-list' -Hover; Collapse-El (Get-El $win 'cboRingType')
  Expand-El (Get-El $win 'cboRingPlan'); Start-Sleep -Milliseconds 600; Shot $win 'rings-plan-list' -Hover
  $li = Find-Role $win 'ListItem' 'Workstation-Rings'; if ($li) { Click-El $li; Start-Sleep -Seconds 1 } else { Collapse-El (Get-El $win 'cboRingPlan') }
  Set-Text (Get-El $win 'txtRingObject') 'Doc Demo Viewer 1.0'
  Invoke-El (Get-El $win 'btnRingExpand'); Start-Sleep -Seconds 3; Shot $win 'rings-expanded'
  Invoke-El (Get-El $win 'btnRingRefreshRuns'); Start-Sleep -Seconds 3; Shot $win 'rings-runs'
}
Section 'tab order' { Invoke-El (Get-El $win 'btnApps'); Focus-Win $win; Click-El (Get-El $win 'btnApps'); 1..16 | ForEach-Object { Send-Keys '{TAB}'; $f = [System.Windows.Automation.AutomationElement]::FocusedElement; Log ('tab -> ' + $f.Current.AutomationId + ' / ' + $f.Current.Name) } }
Section 'light' {
  Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
  foreach ($v in @(@('btnApps', 'light-apps'), @('btnRings', 'light-rings'))) {
    Invoke-El (Get-El $win $v[0]); Start-Sleep -Seconds 2
    # The ring view shows a selected row, which is white text on the accent blue.
    if ($v[0] -eq 'btnRings') { $r = Rows 'dgRingPreview'; if ($r.Count) { Select-El $r[0]; Start-Sleep -Milliseconds 700 } else { Log 'ERROR: the ring preview shows no row to select' } }
    Shot $win $v[1]
  }
  Invoke-El (Get-El $win 'btnOptions'); $d = Find-Last 'Options' 10; Shot $d 'light-options'; Focus-Win $d; Send-Keys '{ESC}'; Start-Sleep -Seconds 1
  Toggle-El (Get-El $win 'toggleTheme')
}
Section 'persist' { Invoke-El (Get-El $win 'btnRings'); Start-Sleep -Milliseconds 600; Invoke-El (Get-El $win 'Close'); Start-Sleep -Seconds 8; $win = Start-App 'start-deploymenthelper.ps1' 'Deployment Helper'; Shot $win 'start-reopen' }
Section 'files' {
  New-Item -ItemType Directory "$OutDir\files" -Force | Out-Null
  Get-ChildItem $appDir -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.FullName.Substring($appDir.Length) -notmatch '^\\(ui|Lib|Module)\\' -and $_.Extension -in '.json', '.jsonl', '.log', '.csv', '.html', '.txt' } | ForEach-Object { Log ('file: ' + $_.FullName.Substring($appDir.Length) + ' ' + $_.Length); Copy-Item $_.FullName ("$OutDir\files\" + $_.Name) -Force }
}
Stop-App 'start-deploymenthelper.ps1'
