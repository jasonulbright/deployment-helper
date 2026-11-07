# Test copy only: make the module path lookup return nothing, as on a computer without the console.
$f = Join-Path $appDir 'Lib\SuiteCommon\SuiteCommon.psm1'
$t = Get-Content -LiteralPath $f -Raw
$patched = [regex]::Replace($t, '(function Resolve-ConfigurationManagerModulePath \{)', '${1} return $null;')
if ($patched -eq $t) { throw 'The module lookup function was not found. The missing-console test cannot run.' }
Set-Content -LiteralPath $f -Value $patched -Encoding UTF8
Remove-Item Env:SMS_ADMIN_UI_PATH -ErrorAction SilentlyContinue
$win = Start-App 'start-deploymenthelper.ps1' 'Deployment Helper'
Size-Win $win 1440 900
$win = Find-Top 'Deployment Helper' 10
Set-Text (Get-El $win 'txtTargetName') 'Doc Demo Viewer 1.0'; Set-Text (Get-El $win 'txtCollection') 'All Workstations'
Invoke-El (Get-El $win 'btnValidate'); Start-Sleep -Seconds 8
Shot $win 'console-missing-validate'
Stop-App 'start-deploymenthelper.ps1'
