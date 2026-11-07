# Troubleshooting

This page lists the most common problems of Deployment Helper with cause and fix. See [Error messages](reference/error-messages.md) for every message.

## What it does

It gives a short fix for each common problem.

## When to use it

Use it when a check, a deployment or a ring run does not work.

## Requirements

None.

## Problems

| Message or symptom | Cause | Fix |
|---|---|---|
| `Failed to connect.` | The site code or the provider is wrong, or the console is missing, or `SMS_ADMIN_UI_PATH` does not point to the `bin\i386` folder of the console. | Check the connection and the console. If the console is installed, set `SMS_ADMIN_UI_PATH` to its `bin\i386` folder and start the tool again. |
| The **Deploy** button is missing. | Not all checks passed. | Select **Validate** and fix the failing checks. |
| The check `Content distributed to DPs` fails. | The content is on no distribution point, or on some of the targeted distribution points only. | Read the logged counts and the distribution status first. Then select the DP groups and **Distribute content** if needed. |
| The check `Collection safe` fails. | The collection is built-in. | Select a custom collection. |
| The check `No duplicate deployment` fails. | The deployment exists. | Select another collection. |
| The first check fails after you change the type. | The object field kept the text of the old type. | Clear the field and select the object with **Browse...**. |
| The ring preview shows no collection. | The seed plan names no collection. | Select **Set collection...**. |
| Promote is disabled. | The previous ring has not reached its threshold. | Select **Refresh and reconcile** later. |
| `Audit log write failed`. | The audit path is not writable. | Set a writable path in **Options > Logging**. |

## Steps

1. Read the log pane. It names the cause of most problems.
2. Read the session log in `Logs`.
3. Do the fix of the table.

## Limits

None.

## Links

- [Error messages](reference/error-messages.md)
- [Files and logs](reference/files-and-logs.md)
