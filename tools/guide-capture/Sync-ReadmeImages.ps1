<#
.SYNOPSIS
    Copies the capture images over the README images that map to them.

.DESCRIPTION
    readme-images.json (next to this script) maps each README image path to the docs\images file that it is a copy of.
    Run the capture first, then this script. It fails when a source image is missing. The capture tool calls it at its end.

.PARAMETER RepoRoot
    The repository folder. The default is the repository that holds this script.

.EXAMPLE
    .\tools\guide-capture\Sync-ReadmeImages.ps1
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
)
$ErrorActionPreference = 'Stop'
$mapPath = Join-Path $PSScriptRoot 'readme-images.json'
$map = Get-Content -LiteralPath $mapPath -Raw | ConvertFrom-Json
$missing = @()
foreach ($entry in $map.PSObject.Properties) {
    $source = Join-Path $RepoRoot ($entry.Value -replace '/', '\')
    if (-not (Test-Path -LiteralPath $source)) { $missing += ('{0} (source of {1})' -f $entry.Value, $entry.Name) }
}
if ($missing.Count) { throw ('A source image is missing: ' + ($missing -join ', ')) }
foreach ($entry in $map.PSObject.Properties) {
    $source = Join-Path $RepoRoot ($entry.Value -replace '/', '\')
    $target = Join-Path $RepoRoot ($entry.Name -replace '/', '\')
    New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
    Copy-Item -LiteralPath $source -Destination $target -Force
    Write-Output ('{0} <- {1}' -f $entry.Name, $entry.Value)
}
