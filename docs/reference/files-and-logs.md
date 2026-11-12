# Files and logs

This page lists every folder and file that Deployment Helper writes.

## What it does

The tool writes in the folder of the script, and in the audit log path and the ring run-state folder that you set. It writes nothing in the registry.

## When to use it

Use it to find a log, a template, a ring plan or the audit trail, or to clean up.

## Requirements

None.

## Files

| Path | Content | Written when |
|---|---|---|
| `DeploymentHelper.prefs.json` | The settings. | You select **OK** in Options. |
| `DeploymentHelper.windowstate.json` | The size, the place, the theme and the last used type. | You close the window. |
| `Logs\DeploymentHelper-<yyyyMMdd-HHmmss>.log` | The log of one session. | During the session. |
| `Logs\DeploymentHelper-startup-<yyyyMMdd-HHmmss>.txt` | The PowerShell transcript of the start. | At the start. |
| `Logs\deployment-audit.jsonl` | The audit trail. One record for each deployment attempt. Append-only. | You confirm a deployment, or a ring is created. |
| `Templates\*.json` | The deployment templates. Four are written at the first start. | You save a template in Options. |
| `Rings\*.json` | The ring plans. Two are written at the first start. | You edit a plan file. |
| `Rings\runs\*.json` | The run state of held ring runs. `closed\` holds finished runs. | A held ring run changes. |
| The file that you save in the Save As dialog (CSV or HTML) | The history export. The dialog starts in `Reports`, and you can choose another folder. | **Export CSV...** or **Export HTML...** in Options. |

The paths of the audit log and of the run-state folder come from the **Logging** page.

## Steps

### Clean up

1. Close the tool.
2. Delete the logs and the reports that you do not need.
3. Keep the audit log. The tool treats it as an immutable record.

## Limits

- The tool never deletes a log, a report or an audit record.
- A deployment is blocked when the audit log cannot be written.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `Audit Log Write Failed` | The audit path is not writable. | Set a writable path in **Options > Logging**. |

## Links

- [Settings](settings.md)
- [Ring deployment](../core/rings.md)
