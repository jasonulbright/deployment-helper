# Changelog

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
