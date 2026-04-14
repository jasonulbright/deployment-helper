# Deployment Helper — Roadmap

**Product:** Deployment Helper (PowerShell/WinForms)
**Repo:** `c:\projects\deploymenthelper\`
**Stack:** PowerShell 5.1, .NET Framework 4.8, WinForms

---

## v1.3.0 — Current (STABLE)

- DP group distribution with search dialogs
- Application + SUG deployment support
- 5-check validation engine with safety guardrails
- Templates (Workstation/Server x Pilot/Production)
- Immutable JSONL audit log
- Dark/light theme
- CSV/HTML history export

---

## v1.4.0 — Functionality Fixes (NEXT)

### Problem
Deployment Helper has UX issues that make it not fully functional for real-world use. These must be fixed before the C# fork (Dployr) begins, so the proven workflow is correct.

### Scope
- Audit the full GUI workflow end-to-end against a live MECM environment
- Identify and fix all UX gaps, broken flows, and unusable features
- Deep research into MECM deployment cmdlets (`New-CMApplicationDeployment`, `New-CMSoftwareUpdateDeployment`, etc.) — parameter validation, verb specificity, edge cases
- Verify every module function works against CM 2509
- Fix deployment creation to match MECM best practices

### Deliverables
- Working deployment workflow: search app → select collection → validate → deploy
- All 20 module functions verified functional
- Updated test coverage
- This version becomes the baseline for Dployr's logic port

---

## v2.0 — Dployr (SEPARATE PRODUCT)

Deployment Helper v1.x remains as the free, open-source PowerShell tool.

The C# Fluent UI evolution lives in a separate private repo as **Dployr** (`c:\projects\dployr\`). See [Dployr ROADMAP](https://github.com/jasonulbright/dployr) for the full vision: drag-drop calendar, ring pipeline, auto-promotion, wave builder, and Packr integration.

Deployment Helper will continue to receive maintenance updates but no major new features. Dployr is the future.
