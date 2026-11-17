# Shared helpers for the Deployment Helper tests of the start script. Dot-source from BeforeAll.

$script:AppRoot = Split-Path -Parent $PSScriptRoot

function Import-AppModule {
    Import-Module (Join-Path $script:AppRoot 'Module\DeploymentHelperCommon.psd1') -Force -DisableNameChecking
}

# The parameter names of one command call: the ones written as -Name, and the keys of a splatted
# hashtable (@params) that the same function or script block builds with a literal or with
# $params['Key'] / $params.Key assignments.
function Get-CommandParameterName {
    param([Parameter(Mandatory)][System.Management.Automation.Language.CommandAst]$Command)
    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($e in $Command.CommandElements) {
        if ($e -is [System.Management.Automation.Language.CommandParameterAst]) { $names.Add($e.ParameterName) }
        elseif ($e -is [System.Management.Automation.Language.VariableExpressionAst] -and $e.Splatted) {
            $var = $e.VariablePath.UserPath
            $scope = $Command.Parent
            while ($scope -and $scope -isnot [System.Management.Automation.Language.FunctionDefinitionAst] -and $scope -isnot [System.Management.Automation.Language.ScriptBlockAst]) { $scope = $scope.Parent }
            if (-not $scope) { continue }
            $assigns = $scope.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)
            foreach ($a in $assigns) {
                $left = $a.Left
                # Only an assignment before the call counts. A new hashtable counts only when its block holds the call,
                # so the other branches of a switch or an if, which build their own hashtable, do not add their keys.
                if ($a.Extent.StartOffset -gt $Command.Extent.StartOffset) { continue }
                if ($left -is [System.Management.Automation.Language.VariableExpressionAst] -and $left.VariablePath.UserPath -eq $var -and $a.Right.Expression -is [System.Management.Automation.Language.HashtableAst]) {
                    $block = $a.Parent
                    while ($block -and $block -isnot [System.Management.Automation.Language.StatementBlockAst] -and $block -isnot [System.Management.Automation.Language.NamedBlockAst]) { $block = $block.Parent }
                    if ($block -and -not ($block.Extent.StartOffset -le $Command.Extent.StartOffset -and $Command.Extent.EndOffset -le $block.Extent.EndOffset)) { continue }
                    foreach ($kv in $a.Right.Expression.KeyValuePairs) { if ($kv.Item1 -is [System.Management.Automation.Language.StringConstantExpressionAst]) { $names.Add($kv.Item1.Value) } }
                }
                elseif ($left -is [System.Management.Automation.Language.IndexExpressionAst] -and $left.Target -is [System.Management.Automation.Language.VariableExpressionAst] -and $left.Target.VariablePath.UserPath -eq $var -and $left.Index -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $names.Add($left.Index.Value)
                }
                elseif ($left -is [System.Management.Automation.Language.MemberExpressionAst] -and $left.Expression -is [System.Management.Automation.Language.VariableExpressionAst] -and $left.Expression.VariablePath.UserPath -eq $var -and $left.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $names.Add($left.Member.Value)
                }
            }
        }
    }
    return @($names | Select-Object -Unique)
}
