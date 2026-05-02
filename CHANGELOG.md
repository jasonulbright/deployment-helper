# Changelog

## [1.0.0] - 2026-05-02

Deployment Helper is a single-pane MECM deployment tool for
Applications, Packages, Task Sequences, and Software Update Groups
with pre-execution validation, safety guardrails, and immutable audit
logging. Ships as a zip + `install.ps1` wrapper; no MSI, no code
signing required.

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
