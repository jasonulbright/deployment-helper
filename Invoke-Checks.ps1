#Requires -Version 5.1
<#
.SYNOPSIS
    Runs the linter and the tests of Deployment Helper.

.DESCRIPTION
    One command for CI and for a test computer:
      1. PSScriptAnalyzer on the module, the start script and the tests. Any finding
         of severity Error or Warning fails the run.
      2. Pester on Tests\. A failed test fails the run. A run that finds no test
         fails too.
    The summary lists the test count as run and the coverage of the module file.
    The coverage is informational. No floor applies until the owner sets one.
    The repository has no window test. Lab and capture scripts live in tools\ and
    never run here.

.PARAMETER SkipLint
    Leave out PSScriptAnalyzer.

.EXAMPLE
    .\Invoke-Checks.ps1
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Console harness: Write-Host is the output channel.')]
[CmdletBinding()]
param(
    [switch]$SkipLint
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$failed = $false

# Pester 5 only: the tests use the 5.x configuration and syntax, and Pester 6 changes them.
Import-Module Pester -MinimumVersion 5.0 -MaximumVersion 5.99.99 -ErrorAction Stop

# ---- 1. Lint ---------------------------------------------------------------
if (-not $SkipLint) {
    Import-Module PSScriptAnalyzer -ErrorAction Stop
    $targets = @(
        Get-ChildItem -LiteralPath (Join-Path $root 'Module') -Include '*.psm1', '*.psd1', '*.ps1' -File -Recurse
        Get-ChildItem -LiteralPath $root -Filter 'start-*.ps1' -File
        Get-ChildItem -LiteralPath (Join-Path $root 'Tests') -Filter '*.ps1' -File -Recurse
        Get-Item -LiteralPath (Join-Path $root 'Invoke-Checks.ps1')
    )
    $settings = Join-Path $root 'Tests\PSScriptAnalyzerSettings.psd1'
    # Test files pass state to mock bodies through a global variable, and they define stub commands in the global scope. Both are intended.
    # Pester sets variables in BeforeAll that the It blocks read. The unused-variable rule reports them.
    $testOnlyExclude = 'PSAvoidGlobalVars', 'PSUseDeclaredVarsMoreThanAssignments', 'PSUseShouldProcessForStateChangingFunctions'
    # The start script keeps its window state in global variables. Event handlers read them.
    # The window asks the user to confirm each destructive step. The functions of the start script take no -WhatIf.
    # The exclusions apply to the start script only. A module function that needs the rule carries its own SuppressMessageAttribute.
    $appExclude = 'PSAvoidGlobalVars', 'PSUseShouldProcessForStateChangingFunctions'
    $findings = @(foreach ($t in $targets) {
        if ($t.FullName -like '*\Tests\*' -or $t.Name -like '*.Tests.ps1') { Invoke-ScriptAnalyzer -Path $t.FullName -Settings $settings -ExcludeRule $testOnlyExclude }
        elseif ($t.Name -like 'start-*.ps1') { Invoke-ScriptAnalyzer -Path $t.FullName -Settings $settings -ExcludeRule $appExclude }
        else { Invoke-ScriptAnalyzer -Path $t.FullName -Settings $settings }
    })
    $blocking = @($findings | Where-Object { $_.Severity -in 'Error', 'Warning' })
    Write-Host ('Lint: {0} files, {1} findings ({2} Error or Warning)' -f $targets.Count, $findings.Count, $blocking.Count)
    foreach ($f in $blocking) { Write-Host ('  {0}:{1} [{2}] {3}' -f (Split-Path $f.ScriptPath -Leaf), $f.Line, $f.RuleName, $f.Message) }
    if ($blocking.Count -gt 0) { $failed = $true }
}

# ---- 2. Tests and coverage -------------------------------------------------
$cfg = New-PesterConfiguration
# Only Tests is searched, so a vendored folder cannot add a test.
$testFiles = @(Get-ChildItem -LiteralPath (Join-Path $root 'Tests') -Recurse -Filter '*.Tests.ps1' -File)
if ($testFiles.Count -eq 0) { Write-Host 'No test file found.'; exit 1 }
$cfg.Run.Path = $testFiles.FullName
$cfg.Run.PassThru = $true
$cfg.Output.Verbosity = 'Minimal'
$cfg.CodeCoverage.Enabled = $true
$cfg.CodeCoverage.OutputPath = Join-Path ([System.IO.Path]::GetTempPath()) 'deploymenthelper-coverage.xml'
$cfg.CodeCoverage.Path = @(Get-ChildItem -LiteralPath (Join-Path $root 'Module') -Filter '*.psm1' -File | ForEach-Object FullName)
$result = Invoke-Pester -Configuration $cfg

$cov = $result.CodeCoverage
$pct = if ($cov -and $cov.CommandsAnalyzedCount -gt 0) { 100 * $cov.CommandsExecutedCount / $cov.CommandsAnalyzedCount } else { 0 }
Write-Host ('Tests: {0} run, {1} passed, {2} failed, {3} skipped' -f $result.TotalCount, $result.PassedCount, $result.FailedCount, $result.SkippedCount)
Write-Host ('Coverage: {0:N1}% of {1} module commands (no floor)' -f $pct, $cov.CommandsAnalyzedCount)

if ($result.TotalCount -eq 0) { Write-Host 'No test ran.'; $failed = $true }
if ($result.FailedCount -gt 0) { $failed = $true }
if ($result.Result -ne 'Passed') { Write-Host ('Pester result: {0}' -f $result.Result); $failed = $true }

if ($failed) { Write-Host 'CHECKS FAILED'; exit 1 }
Write-Host 'CHECKS PASSED'
exit 0
