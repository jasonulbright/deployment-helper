# Settings

This page lists what Deployment Helper saves and where.

## What it does

The tool saves its settings in `DeploymentHelper.prefs.json` in the folder of the script. The Options dialog writes the file when you select **OK**. The window state is in `DeploymentHelper.windowstate.json`.

## When to use it

Use it to copy the settings to another computer or to reset them.

## Requirements

None.

## Settings

| Key | Option | Default |
|---|---|---|
| `SiteCode` | **Site code** | Empty. |
| `SMSProvider` | **SMS Provider** | Empty. |
| `DeploymentAuditLogPath` | **Deployment audit log (JSONL)** | `Logs\deployment-audit.jsonl` beside the script. |
| `RingRunStatePath` | **Ring run-state folder** | `Rings\runs` beside the script. |

`DeploymentHelper.windowstate.json` keeps the size, the place, the theme and the last used deployment type. A non-empty command line parameter `-SiteCode` or `-SMSProvider` overrides the saved value at the start. When you then select **OK** in Options, the tool saves the current values, including the overrides, to `DeploymentHelper.prefs.json`.

## Steps

### Reset the settings

1. Close the tool.
2. Delete `DeploymentHelper.prefs.json` and `DeploymentHelper.windowstate.json`.
3. Start the tool. The window of the first start shows.

## Limits

- The files are in the folder of the script. Extract the tool to a folder that you can write to.
- Templates and ring plans are separate files. See [Files and logs](files-and-logs.md).

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| The settings do not stick. | The folder is not writable. | Move the tool to a writable folder. |

## Links

- [Options dialog](../core/options.md)
- [Files and logs](files-and-logs.md)
