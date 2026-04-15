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

## v1.4.0 — Functionality Fixes (COMPLETE)

- E2E workflow validated against live MECM environment
- 51 Pester 5 tests (all passing)
- MW parameter fix, form reset, duplicate check
- All 20 module functions verified functional against CM 2509
- This version is the baseline for Dployr's logic port

---

## v2.0 — Dployr (SEPARATE PRODUCT)

Deployment Helper v1.x remains as the free, open-source PowerShell tool.

The C# Fluent UI evolution lives in a separate private repo as **Dployr** (`c:\projects\dployr\`). See [Dployr ROADMAP](https://github.com/jasonulbright/dployr) for the full vision: drag-drop calendar, ring pipeline, auto-promotion, wave builder, and Packr integration.

Deployment Helper will continue to receive maintenance updates but no major new features. Dployr is the future.
