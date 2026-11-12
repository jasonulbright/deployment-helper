# Permissions

This page lists the rights that Deployment Helper needs.

## What it does

The checks and the lists read the site. **Deploy**, **Distribute Content**, **Create Deployments** and Promote change the site with the rights of the account that started the tool.

## When to use it

Use it to set up an account, or to find out why a read or a deployment is refused.

## Requirements

| Action | Right |
|---|---|
| Start the tool | Read and write in the folder of the script. |
| Browse, Validate | The Configuration Manager console on the computer, and read access to the SMS Provider under your Configuration Manager role. |
| **Deploy** of an application, a package, a task sequence or a software update group | The right to create a deployment of that type for the collection. |
| **Distribute Content** | The right to distribute content to the distribution point groups. |
| Ring runs and Promote | The rights of the deployment type for each ring collection. A ring with a `DPGroup` also needs the right to distribute content to that group. |
| The audit log and the run-state folder | Write access to the path. A UNC run-state folder needs write access to the share. |

## Steps

1. Start the tool as the account that has the Configuration Manager role.
2. Select **Validate** to test the read access.

## Limits

- The tool blocks every collection whose ID starts with `SMS`, whatever the rights.
- The tool does not check the rights before a deployment. A refused deployment shows in the log as `deployment failed`.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `Failed to connect.` | The account has no access to the provider, or the console is missing. | Check the role, the provider name and the console. |
| `... deployment failed:` | The role lacks a right. | Use an account with the right. |

## Links

- [Network](network.md)
- [Application deployment](../core/apps.md)
