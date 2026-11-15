#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Headless guard for the images of the README: every image is a guide image or comes from the capture output.
# It parses README.md and compares files. It opens no window and starts no capture.

BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:Readme = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'README.md') -Raw
    $script:Map = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'tools\guide-capture\readme-images.json') -Raw | ConvertFrom-Json
    $script:MapNames = @($script:Map.PSObject.Properties | ForEach-Object { $_.Name })
    $script:ReadmeImages = @([regex]::Matches($script:Readme, '!\[[^\]]*\]\(([^)\s]+\.(?:png|gif|jpg|jpeg))\)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
}

BeforeDiscovery {
    $root = Split-Path -Parent $PSScriptRoot
    $map = Get-Content -LiteralPath (Join-Path $root 'tools\guide-capture\readme-images.json') -Raw | ConvertFrom-Json
    $script:MapCases = @($map.PSObject.Properties | ForEach-Object { @{ Target = $_.Name; Source = $_.Value } })
    $readme = Get-Content -LiteralPath (Join-Path $root 'README.md') -Raw
    $script:ImageCases = @([regex]::Matches($readme, '!\[[^\]]*\]\(([^)\s]+\.(?:png|gif|jpg|jpeg))\)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique | ForEach-Object { @{ Image = $_ } })
}

Describe 'README images' {
    It 'shows at least one image' {
        $script:ReadmeImages.Count | Should -BeGreaterThan 0
    }

    It '<Image> is a guide image or is listed in readme-images.json' -ForEach $script:ImageCases {
        ($Image -like 'docs/images/*' -or $script:MapNames -contains $Image) | Should -BeTrue
    }

    It '<Target> is byte-identical to its source <Source>' -ForEach $script:MapCases {
        $t = Join-Path $script:RepoRoot ($Target -replace '/', '\')
        $s = Join-Path $script:RepoRoot ($Source -replace '/', '\')
        Test-Path -LiteralPath $s | Should -BeTrue -Because "the source $Source must exist"
        Test-Path -LiteralPath $t | Should -BeTrue -Because "the README image $Target must exist"
        (Get-FileHash -LiteralPath $t -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $s -Algorithm SHA256).Hash -Because 'run tools\guide-capture\Sync-ReadmeImages.ps1'
    }

    It 'maps every README image to a docs\images file' {
        foreach ($p in $script:Map.PSObject.Properties) { $p.Value | Should -BeLike 'docs/images/*' }
    }
}
