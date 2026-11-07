# Glossary

This page defines the terms of Deployment Helper.

## What it does

It explains the terms that the tool and this guide use.

## When to use it

Use it when a term is not clear.

## Requirements

None.

## Terms

| Term | Meaning |
|---|---|
| Deployment | The assignment of an object to a collection on the site. |
| Purpose | **Available** (the user can install) or **Required** (the client installs by the deadline). |
| Deadline | The time by which a Required deployment installs. |
| Check | One of the five tests before a deployment. |
| Audit log | The append-only JSONL file with one record for each deployment attempt. |
| Template | A saved set of form values in `Templates`. |
| Ring plan | An ordered list of rings in `Rings\<plan>.json`. |
| Ring | One collection of a plan with its dates and options. |
| Run | One use of a plan for one object. |
| Held ring | A ring that the tool has not created yet. Promote creates it. |
| Promote | Creates the next held ring after the checks and the threshold. |
| Reconcile | Reads the site to update the state of a run. |
| DP group | A group of distribution points. |
| Built-in collection | A collection whose ID starts with `SMS`. The tool blocks it. |
| SMS Provider | The server role that gives access to the site data. |

## Steps

None.

## Limits

None.

## Troubleshooting

None.

## Links

- [Application deployment](../core/apps.md)
