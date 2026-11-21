# App-specific names for the window tests.
$script:AppConfig = @{
    Title           = 'Deployment Helper'
    Shell           = 'start-deploymenthelper.ps1'
    Xaml            = 'MainWindow.xaml'
    LogLabel        = 'lblLogOutput'
    WorkflowButtons = @('btnApps', 'btnPackages', 'btnTaskSequences', 'btnSUG', 'btnRings', 'btnRollout')
    OptionsButtons  = @('btnOptions')
    FormPane        = 'formPane'
    RingPane        = 'ringPane'
    RolloutPane     = 'rolloutPane'
}
