# Ring plan files from Update Planner

Update Planner writes its plan as a Suite ring plan. Deployment Helper loads the file as a ring plan with no change. Update Planner (**Send Plan**) saves the plan in the handoff folder of the site, `sites\<SITE>\handoff\ring-plan.json`, and Deployment Helper offers it in the Ring Deployment view with **Import Plan**. You do not copy the file by hand. The view shows the tool, its version, the site and the age of the plan, and marks a plan that is older than the stale limit (Options, page **Imports**). **Planned Updates...** shows the `Updates` list of the selected plan. The list is read only.

The file has the fields of a Deployment Helper ring plan (`Name`, `Description`, `HoldLaterRings`, `TimeBasedOn`, and `Rings[]` with the ring fields). It also has extra fields that Deployment Helper ignores:

- `Format` (`SuiteRingPlan`) and `SchemaVersion` (`1`).
- `UpdatePlanner`: the plan ID, the site, the parameters, the summary and the skipped updates.
- `Producer` (`Name` and `Version` of the tool), `CreatedUtc` (`yyyy-MM-ddTHH:mm:ssZ`) and `Site` (`SiteCode`): the header of a handoff file. The tool that reads the file shows them. A file with no valid `CreatedUtc` or with a `SchemaVersion` above 1 is refused.
- In each ring: `MemberCount`, `MaintenanceWindows`, `StartDate`, `DeadlineDate` and `Updates` (the planned updates of the wave).

Deployment Helper shows the `Updates` list read only and does not use it to deploy. A ring deploys the software update group that you choose in Deployment Helper.

The schema is [schema/ring-plan.schema.json](schema/ring-plan.schema.json). The test `Tests/RingPlanContract.Tests.ps1` loads the shared fixture `Tests/Fixtures/shared-ring-plan.json` with `Import-RingPlan`. The same fixture and schema are in the Update Planner repository (`docs/schema`, `Tests/Fixtures`). Both tests check the SHA-256 of the two files, so a change of one copy fails the test of the other repository until both change.

Rules that Deployment Helper checks (`Test-RingPlan`): each ring deadline offset is later than its available offset and later than the offset of the ring before. Update Planner needs **Days to deadline** of 1 or more and **Days between waves** of 1 or more to write such a plan.

The schema checks the shape of the file only. Import-RingPlan also checks the collection IDs: it refuses a built-in ID such as SMS00001 and an ID that two rings share. A file that passes the schema can still fail to load.
