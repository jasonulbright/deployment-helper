# Changelog

## [1.0.1] - 2026-04-24

- Sidebar pressed-state fix: replaced the MahApps `Square` button template with a custom `SidebarButtonTemplate` that shade-lifts to `#3A3A3A` on press instead of inverting colors via the default Visual State Manager. The default template made white button text invisible during the brief press flash on the dark sidebar; the new template keeps `Foreground=White` constant and only swaps the background. Applied to both `WorkflowButtonStyle` and `OptionsButtonStyle`.
- Tooltips added to the workflow form: target/collection Browse buttons, DP-group picker, distribute content, validate, and deploy actions, plus the dark/light theme toggle. Each tooltip is a one-line description of what the button does.
- **OPTIONS** sidebar button retitled to **Options** (title-case) to match the rest of the sidebar; previous all-caps was a holdover from earlier brand iteration.

## [1.0.0] - 2026-04-22

Initial release.
