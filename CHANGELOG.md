# Changelog

All notable changes to the Deployment Helper are documented in this file.

## [1.1.0] - 2026-02-27

### Added

- **Software Update Group (SUG) deployment support**
  - Deployment type selector (Application / Software Update Group) in form
  - `Test-SUGExists` validation using `Get-CMSoftwareUpdateGroup`
  - `Invoke-SUGDeployment` execution using `New-CMSoftwareUpdateDeployment`
  - Required SUGs auto-set download fallback: allow from default site boundary group DP and unprotected DP
  - Validate/deploy handlers dispatch between Application and SUG based on type selector
  - SUG validation skips content distribution and duplicate deployment checks (not applicable)

- **Metered connection checkbox**
  - "Allow download past deadline (metered connections)" checkbox in deployment form
  - Auto-checked when Required purpose is selected (both Application and SUG)
  - Passed to `Invoke-ApplicationDeployment` (`-AllowMeteredConnection`) and `Invoke-SUGDeployment` (`-AllowUseMeteredNetwork`)

- **Save Template from GUI**
  - Save Template button alongside Validate and Deploy
  - Prompts for template name, saves current form configuration as JSON
  - Reloads template ComboBox after save
  - `Save-DeploymentTemplate` module function

- **Deployment log includes deployment type**
  - JSONL records now include `DeploymentType` field (`Application` or `SUG`)

### Changed

- `Get-DeploymentPreview` accepts `$TargetObject` + `$DeploymentType` instead of `$Application` (supports both app and SUG objects)
- Module manifest exports expanded from 17 to 20 functions

### Fixed

- Window height increased 100px (820 -> 920) to prevent form overlap
- Reboot/MW checkboxes no longer overlap validate/deploy buttons (form panel height 340 -> 430, row positions adjusted)

## [1.0.0] - 2026-02-27

### Added

- **GUI application** (`start-deploymenthelper.ps1`) with WinForms interface
  - Header panel with title and subtitle
  - Connection bar with site code, SMS provider, status, and Connect button
  - 10-row deployment form: change ticket, application, collection, template, purpose, available date, deadline, notification, maintenance window overrides, validate/deploy buttons
  - Validation results panel with colored `[PASS]`/`[FAIL]`/`[INFO]` output
  - Live log console with timestamped progress messages
  - Status bar with connection and deployment status

- **Pre-execution validation engine** (5 sequential checks)
  - Application exists in MECM
  - Content fully distributed to all targeted DPs
  - Collection exists and is a Device collection
  - Collection is not a built-in system collection (blocks all `SMS000*` IDs)
  - No duplicate deployment already exists for this app/collection combination

- **Safety guardrails**
  - Built-in system collections blocked (SMS00001, SMS00004, SMS000C1, and all `SMS000*` pattern)
  - Deploy button disabled until all 5 validation checks pass
  - Confirmation dialog before execution showing app, version, collection, device count, purpose

- **Deployment templates** (4 placeholders)
  - Workstation Pilot, Workstation Production, Server Pilot, Server Production
  - Template selection auto-populates purpose, notification, maintenance window, deadline offset
  - JSON schema for user-defined templates

- **Immutable deployment audit log** (JSONL format)
  - One JSON object per line, append-only
  - Records: timestamp, user, change ticket, app name/version, collection, member count, purpose, deadline, deployment ID, result
  - Both success and failure outcomes logged
  - Configurable log path via Preferences (default: local `Logs\deployment-log.jsonl`)

- **Export**
  - CSV export of deployment history
  - HTML export with styled report, color-coded Result column (green/red)

- **Dark mode** with full theme support
  - Custom `DarkToolStripRenderer` for MenuStrip and StatusStrip
  - Configurable via File > Preferences
  - Persisted in `DeploymentHelper.prefs.json`

- **Menu bar** with File (Preferences, Exit) and Help (About)

- **Window state persistence** across sessions (`DeploymentHelper.windowstate.json`)

- **Core module** (`DeploymentHelperCommon.psm1`) with 17 exported functions
  - CM site connection management
  - 5 validation functions + deployment preview
  - Deployment execution via `New-CMApplicationDeployment`
  - JSONL deployment log (write + read)
  - Template loading
  - CSV and HTML export
