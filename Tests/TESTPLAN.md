# Deployment Helper: test plan

Measured 2026-11-17 with `Invoke-Checks.ps1` (Pester 5.7.1, PSScriptAnalyzer 1.25.0) on a lab computer, with the window tests off. Counts are tests as run; each data row counts as one test.

| | Windows PowerShell 5.1 | PowerShell 7 |
|---|---|---|
| Passed | 788 | 778 |
| Skipped (window tests, opt-in off) | 101 | 100 |

Run: `.\Invoke-Checks.ps1` (lint, then tests and coverage of the module file; the coverage is informational, no floor applies). It runs in CI on both hosts (`.github/workflows/checks.yml`).

Window tests: the `*.Window.Tests.ps1` files open windows only when `SUITE_ALLOW_WINDOW_TESTS` is `1`. The lab run sets it; CI does not. Otherwise Pester reports those tests as Skipped with that reason, not as Failed. `WindowGuard.Tests.ps1` tests the guard (`Test-WindowTestsAllowed` in `Tests/WindowHarness.ps1`). `WindowCleanup.Guard.Tests.ps1` checks that a window test leaves no window open. Both open no window. `ButtonStates.Window.Tests.ps1` needs the real mouse and runs only in a visible run (`SUITE_VISIBLE_WINDOWS=1`).

## What is covered

| Area | File | Contents |
|---|---|---|
| Module | `DeploymentHelperCommon.Tests.ps1` | The module functions with stubbed ConfigMgr commands (`CMStubs.ps1`) |
| Ring plans | `RingDeployment.Tests.ps1` | Ring plans, ring runs and ring deployment |
| Plan contract | `RingPlanContract.Tests.ps1` | The ring plan that Update Planner writes, read from `Fixtures` |
| Start script | `Shell.Logic.Tests.ps1`, `Script.Contract.Tests.ps1`, `Startup.Tests.ps1` | The functions of the start script cut from its syntax tree with plain objects for the controls, every named parameter of a call into the module and SuiteCommon, and start-up without the ConfigMgr console |
| Window layout | `Window.Tests.ps1`, `WindowXaml.Tests.ps1` | The main window and every dialog load, every control that the script looks up exists, no clipped text and no control outside the window in both panes, both themes, two sizes |
| Windows | `Chrome.Window.Tests.ps1`, `Contrast.Window.Tests.ps1` | Chrome and buttons of the main window and every dialog, the disabled state of the controls, both themes |
| Density | `Density.Tests.ps1`, `Density.Window.Tests.ps1` | The compact 1080p layout from the XAML, and measured on screen with vacant-space checks |
| Button states | `ButtonStates.Window.Tests.ps1` | Hover, focus and pressed fills of the main window buttons, both themes (visible run only) |
| Window guards | `WindowGuard.Tests.ps1`, `WindowCleanup.Guard.Tests.ps1` | The opt-in guard, its use in every window test file, and window cleanup |
| Theme guards | `Theme.Guard.Tests.ps1`, `Theme.Source.Guard.Tests.ps1`, `ButtonStyle.Guard.Tests.ps1`, `FontWeight.Guard.Tests.ps1`, `Footer.Guard.Tests.ps1` | White text on blue, fonts independent of the theme, banned words, one themed button style (also in the vendored SuiteCommon dialogs), Title Case, the footer |
| Encoding | `Encoding.Guard.Tests.ps1` | Plain ASCII in every PowerShell file and every file that PowerShell reads |
| Documents and images | `Documentation.Tests.ps1`, `Readme.Images.Guard.Tests.ps1`, `GuideImages.Guard.Tests.ps1` | Index and links, every README image is a guide image, every guide image was produced by a capture run |
| Capture tool | `GuideCapture.Tests.ps1` | The capture scripts parse, take the lab lock and clean up after themselves |

## Not covered, and why

- The tests stub the ConfigMgr commands. No test calls a real site.
- The deploy handlers run inside the window with a site. The tests cover the module calls that they make, not the clicks.
- The window handlers (a click on Validate or Deploy) run inside the window with a site. The shell logic tests cover the helper functions that they call, not the clicks.
- The settings save in SuiteCommon is not atomic. The fix belongs in `app-packager-suite`.