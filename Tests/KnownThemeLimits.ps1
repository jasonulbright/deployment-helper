# The known limits of the theme that the app cannot fix itself. Theme.Guard.Tests.ps1 reads this list.
# Each entry needs an Id, the place (Where), the reason that says why the limit stays (Reason) and its owner (Source).
# A limit stays inert while the condition in its test holds.

$script:KnownThemeLimit = @(
    @{
        Id     = 'sidebar-disabled-white-on-blue'
        Where  = 'The sidebar buttons (Apps, Packages, Task Sequences, Software Update Groups, Ring Deployment, Options) in the light theme'
        Reason = 'SuiteCommon (Set-ThemeForeground) sets a local white foreground on a blue button in the light theme. A local value beats the gray text of the disabled state in the button style, so a disabled sidebar button would keep white text on blue and look enabled. The app never disables a sidebar button, so the state does not occur.'
        Source = 'Lib\SuiteCommon\SuiteCommon.psm1, Set-ThemeForeground. The owner of the library is app-packager-suite. The fix belongs there.'
    }
)
