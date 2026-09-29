# Changelog

## [2026.09.29.0012] - 2026-09-29

## Browse lists show 11 of 11 lab applications instead of 1 row

### Fixed

- Show every object in the browse lists instead of one row of list data.
- Show plain, full column names in the browse dialogs.
- Fit the Ring Deployment view in 1320 x 860 without a scroll bar.
- Leave a gap between the scroll bar and the buttons beside it.
- Draw an outline around both ring grids.
- Keep the Note and Promote columns visible in the open-runs grid.
- Align the Packages network and rerun rows with the other field labels.
- Fit the template toolbar buttons above the template list.

### Changed

- Call the product Configuration Manager in every label and tooltip.
- Hide the deadline field for Available deployments instead of disabling it.
- Use the same checkbox names for every deployment type and in templates.
- Name the collection ID column the same in both ring grids.
- Use one muted text color, one button height, and one margin in every dialog.
- Open at 1320 x 860 by default with shorter ring grids and log drawer.
- Update the README screenshots.
- Update the shared SuiteCommon module to 2026.09.27.0037.

## [2026.09.27.0011] - 2026-09-27

## The ring view shows the preview and open runs together at 1320 x 820

### Fixed

- Show the CI ID in the application browse list instead of an always-empty package ID.
- Remove the always-empty boot image column from the task sequence browse list.
- Widen the ring grid columns so no header is cut off.
- Show full run labels, notes, and checks as tooltips in the ring grids.

### Changed

- Move the ring preview buttons to the preview header line.
- Update the README screenshots and add one of the ring view.

## [2026.09.26.0010] - 2026-09-26

## One confirmation creates the deployments for all 4 rings of a plan

### Added

- Add a Ring deployment view that deploys one object through a ring plan.
- Seed two ring plans, Workstation-Rings and Server-Rings, with no target collections.
- Create every ring at once with future available times and deadlines.
- Hold later rings on request; Promote creates the next ring.
- Disable Promote until the previous ring reaches its success threshold.
- Show targeted, success, error, and in-progress counts for each created ring.
- Mark deleted ring deployments Removed and close finished runs on Reconcile.
- Refuse Promote when another session changed or locked the run file.
- Add a ring run-state folder setting; a UNC path shares runs with a team.
- Add plan name, ring index, ring name, and run ID to ring audit records.
- Recover an interrupted ring create on the next Refresh and reconcile.
- Stop the ring run when an audit record fails after a create.

### Fixed

- Set Required package and task sequence deadlines as a schedule, not an expiry.
- Block every built-in collection by the SMS ID prefix, including SMSDM collections.
- Block a deployment when the target collection has no collection ID.
- Keep ring columns in the CSV export when earlier records do not have them.
- Check deadlines against UTC time when the time basis is UTC.
- Detect an existing software update group deployment in the duplicate check.
- Match names that contain brackets or other wildcard characters literally.
- Block a deployment when the duplicate check cannot run.
- Run the duplicate check again right before each deployment is created.
- Block a deployment when the audit log cannot be written.
- Show the deployment ID when its audit record cannot be written.

### Changed

- Update the shared SuiteCommon module to 2026.09.25.0036.

## [2026.09.25.0009] - 2026-09-25

## Browse opens 5 object types with no search term

### Changed

- Update the shared SuiteCommon module to 2026.09.25.0033.
- List every application, package, task sequence, or update group when Browse opens.
- Filter the browse list as you type; no search term is required.
- Load browse lists in the background with a progress dialog and Cancel.
- Keep browse lists for the session; Refresh, or Shift+Browse for collections, reloads.
- Browse device collections in their console folder tree.

## [2026.09.21.0008] - 2026-09-21

### Fixed

- Use the site code and provider from the suite launcher when the tool has none saved.
- Show the site code and provider in use in the startup log line.

### Changed

- Update the README screenshot.

## [2026.09.21.0007] - 2026-09-21

### Changed

- Use the product name Configuration Manager in the application and the documentation.
- Use date versions.
- Update the shared SuiteCommon module to 2026.09.21.0031.

## [1.2.3] - 2026-09-04

### Fixed

- **Version labels read the script header.** The sidebar version and the
  About panel carried literal version strings that no release updated;
  both now read the entry script's `Version` header at startup, so the
  window always names the version that is actually installed.

## [1.2.2] - 2026-09-04

### Changed

- **Vendored `SuiteCommon` 0.4.3.** The module repairs the process
  PSModulePath at import: a Windows PowerShell process launched from
  PowerShell 7 inherits the 7.x module directories, and the background
  runspace opened later autoloaded a Microsoft.PowerShell.Utility without
  Get-FileHash or ConvertFrom-Json, so background operations failed with
  an unrelated "term not recognized". A background runspace whose module
  import fails is disposed and the original error thrown instead of being
  returned as an opened but unusable worker.

## [1.2.1] - 2026-08-16

### Changed

- **Vendored `SuiteCommon` 0.3.2.** Window restore applies the saved
  geometry before maximizing, so un-maximizing returns to the saved size
  instead of the XAML defaults.

## [1.2.0] - 2026-08-16

### Changed

- **Window chrome, theming, and the message dialog now come from the
  vendored `SuiteCommon` module** (0.3.0): the title-bar drag block,
  action-button theming, window-state persistence, and
  `Show-ThemedMessage` load from `Lib\SuiteCommon\`. Dialog buttons
  standardize at the suite's 32px height (previously 30px here).
  Behavior gains: hook state no longer leaks on window close, a
  maximized close persists the pre-maximize geometry, an off-screen
  saved position clamps into the nearest monitor, the dialog's inactive
  title-bar brushes now copy from the owner's inactive properties, and
  Escape closes OK-only dialogs.

## [1.1.0] - 2026-08-14

### Changed

- **Shared plumbing moved to the vendored `SuiteCommon` module.** Logging
  (`Initialize-Logging`, `Write-Log`) and CM site connection
  (`Connect-CMSite`, `Disconnect-CMSite`, `Test-CMConnection`) now load
  from `Lib\SuiteCommon\`, shared across the tool suite and synced from
  the suite-core repository instead of hand-edited per repo. The
  connection additionally gains behavior this tool's own copy lacked: a
  globally scoped CMSite PSDrive, normalized ConfigurationManager module
  path resolution with known-install-path fallback, provider rebind when
  the configured SMS Provider changes, and rebuild of a stale drive whose
  provider connection died. `Initialize-Logging` gains `-Attach`.

## [1.0.0] - 2026-05-02

Deployment Helper is a single-pane MECM deployment tool for
Applications, Packages, Task Sequences, and Software Update Groups
with pre-execution validation, safety guardrails, and immutable audit
logging. Extract the zip and run `start-deploymenthelper.ps1`.

### Features

- **Sidebar navigation** across the four target types (Apps,
  Packages, Task Sequences, Software Update Groups) plus an Options
  modal. Theme toggle bottom-docked on the sidebar.
- **Search dialogs** for target object and target collection — filtered
  DataGrid results, pick-and-paste into the workflow form.
- **DP Group Picker** modal — pick one or more DP groups for content
  distribution, no need to remember exact group names.
- **Pre-flight validation** — confirms the target exists in MECM, the
  collection is a device collection (not user, not `SMS000*`), no
  duplicate deployment exists, and content has been distributed to at
  least one DP. Validation runs before any `New-CM*Deployment` call.
- **Available + Deadline date pickers** with Local-time / UTC toggle.
- **Notification mode picker** (Display All / Hide notifications and
  restarts / etc.).
- **Distribute content** action — runs `Start-CMContentDistribution`
  against the selected DP groups in one click.
- **Deploy templates** — save the current form state as a named
  template (target type, collection, purpose, dates, notification);
  reload by clicking Apply Template. Templates persist across
  sessions.
- **MahApps Dark.Steel / Light.Blue themes** with live swap.
- **Title-bar drag fallback** — native `WM_NCHITTEST` hook + managed
  `DragMove` for the main window and every modal dialog so the title
  bar drags reliably under any host.
- **Immutable audit log** — every deployment writes a JSON record
  (target, collection, purpose, deployer, timestamp) to a per-day
  log file under `Logs/`.
- **Window state persistence** — size, position, last-active target
  type restored across launches.

### Stack

- PowerShell 5.1 + .NET Framework 4.7.2+
- WPF + MahApps.Metro (vendored DLLs in `Lib/`)
- ConfigurationManager PowerShell module (provided by the MECM
  Console install)
