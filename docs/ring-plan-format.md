# Ring plan files from Update Planner

Update Planner writes its plan as a Suite ring plan. Deployment Helper loads the file as a ring plan with no change. Put the file in `Rings\` or open it with the ring plan picker.

The file has the fields of a Deployment Helper ring plan (`Name`, `Description`, `HoldLaterRings`, `TimeBasedOn`, and `Rings[]` with the ring fields). It also has extra fields that Deployment Helper ignores:

- `Format` (`SuiteRingPlan`) and `SchemaVersion` (`1`).
- `UpdatePlanner`: the plan ID, the site, the parameters, the summary and the skipped updates.
- In each ring: `MemberCount`, `MaintenanceWindows`, `StartDate`, `DeadlineDate` and `Updates` (the planned updates of the wave).

Deployment Helper does not use the `Updates` list. A ring deploys the software update group that you choose in Deployment Helper.

The schema is [schema/ring-plan.schema.json](schema/ring-plan.schema.json). The test `Tests/RingPlanContract.Tests.ps1` loads the shared fixture `Tests/Fixtures/shared-ring-plan.json` with `Import-RingPlan`. The same fixture and schema are in the Update Planner repository (`docs/schema`, `Tests/Fixtures`). Both tests check the SHA-256 of the two files, so a change of one copy fails the test of the other repository until both change.

Rules that Deployment Helper checks (`Test-RingPlan`): each ring deadline offset is later than its available offset and later than the offset of the ring before. Update Planner needs **Days to deadline** of 1 or more and **Days between waves** of 1 or more to write such a plan.

The schema checks the shape of the file only. Import-RingPlan also checks the collection IDs: it refuses a built-in ID such as SMS00001 and an ID that two rings share. A file that passes the schema can still fail to load.
