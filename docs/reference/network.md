# Network

This page lists every outbound call of Deployment Helper.

## What it does

The tool talks to the SMS Provider through the console module. When it is installed by the suite installer, it also checks GitHub once a day for a newer suite version. It sends no user data and no telemetry.

## When to use it

Use it to open the firewall for the tool, or to check the traffic.

## Requirements

| Call | Target | When | Content |
|---|---|---|---|
| ConfigurationManager module | The SMS Provider in **Options > Connection**, through the site drive that the console module opens. The ports are those of the console. | Browse, Validate, **Distribute content**, **Deploy** and the ring actions. | `Get-CM*` reads and the deployment commands. |
| UNC share | The run-state folder, if you set a UNC path. | Ring runs. | The run files. |
| `https://api.github.com/repos/<repository>/releases/latest` | HTTPS to GitHub. | At the start of the window, once a day, only when the tool is installed by the suite installer. | The check for a newer suite version. |

## Steps

None.

## Limits

- A plain extract of the zip (no `suite-manifest.json`) makes no call to GitHub. A failed update check is ignored. The tool does not change a proxy setting.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `Failed to connect.` | The provider is not reachable. | Check the name, the firewall and the role. |

## Links

- [Permissions](permissions.md)
- [Options dialog](../core/options.md)
