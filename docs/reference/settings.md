# Settings

This page lists what Deployment Helper saves and where.

## What it does

The tool saves its settings in `DeploymentHelper.prefs.json` in the folder of the script. When the AppPackager Suite launcher chose a site, the tool saves `SiteCode` and `SMSProvider` for that site in `%LOCALAPPDATA%\AppPackagerData\sites\<SITE>\tools\deployment-helper\prefs.json`, and the other keys stay in the file beside the script. The Options dialog writes the files when you select **OK**. The window state is in `DeploymentHelper.windowstate.json`.

## When to use it

Use it to copy the settings to another computer or to reset them.

## Requirements

None.

## Settings

| Key | Option | Default |
|---|---|---|
| `SiteCode` | **Site Code** | Empty. The launcher chooses it. It is read-only in Options when a site is chosen. |
| `SMSProvider` | **SMS Provider** | Empty. |
| `DeploymentAuditLogPath` | **Deployment Audit Log (JSONL)** | Empty. Empty means `deployment-audit.jsonl` in the site data folder. With no site it means `Logs\deployment-audit.jsonl` beside the script. |
| `RingRunStatePath` | **Ring Run-State Folder** | Empty. Empty means `rings\runs` in the site data folder. With no site it means `Rings\runs` beside the script. |
| `HandoffStaleValue`, `HandoffStaleUnit` | **Imports** page: the age after which the plan from Update Planner and the Check ConfigMgr result are marked STALE. | `3` and `days`. The unit is `hours` or `days`. |
| `RolloutRunPath` | The folder of the update rollout run files. | Empty. Empty means `rollouts` in the site data folder. With no site it means `Rollouts` beside the script. |
| `RolloutNameTemplate`, `RolloutFolder` | **Name Template** and **Folder** of the Update Rollout view. | `{App} - {Role} - Installed - Version LT {Version}`, empty. |
| `RolloutLimitingServer`, `RolloutLimitingWorkstation`, `RolloutLimitingTest`, `RolloutTestCollection` | The limiting collections and the test collection of the Update Rollout view. | `All Systems` for the three limits, empty test collection. |
| `RolloutRefreshDays`, `RolloutAvailableDelayHours`, `RolloutTestDeadlineHours`, `RolloutProdDeadlineHours`, `RolloutNotification`, `RolloutDefaultRole` | The refresh, the times, the notification and the default role of the Update Rollout view. | 1, 0, 24, 72, `DisplayAll`, `Both`. |

A path key holds a path only when you typed one. An older version saved the default path in the file. The tool treats that old default as empty. A path that you chose stays.

`DeploymentHelper.windowstate.json` keeps the size, the place, the theme and the last used deployment type. A non-empty command line parameter `-SMSProvider` overrides the saved value at the start. `-SiteCode` does the same when no site is chosen. With a site, the launcher decides and the tool ignores `-SiteCode`. When you then select **OK** in Options, the tool saves the current values, including the overrides, to `DeploymentHelper.prefs.json`.

## Sites

The launcher chooses the site. A tool that is open keeps its site until you restart it. The first start after an update copies the old data into `%LOCALAPPDATA%\AppPackagerData\sites\<SITE>\tools\deployment-helper` and keeps the old files. See [Options dialog](../core/options.md#sites).

## Steps

### Reset the settings

1. Close the tool.
2. Delete `DeploymentHelper.prefs.json` and `DeploymentHelper.windowstate.json`.
3. Start the tool. The window of the first start shows.

## Limits

- The global files are in the folder of the script. Extract the tool to a folder that you can write to.
- Templates and ring plans are separate files. See [Files and logs](files-and-logs.md).

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| The settings do not stick. | The folder is not writable. | Move the tool to a writable folder. |

## Links

- [Options dialog](../core/options.md)
- [Files and logs](files-and-logs.md)
