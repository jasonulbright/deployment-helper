#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }

<#
.SYNOPSIS
    Pester 5.x test for the compact 1080p density of the main window (XAML only, no window is opened).
    The shared style values are asserted in one place. The window test Density.Window.Tests.ps1 measures them on screen.
#>

BeforeAll {
    $script:Xaml = [xml](Get-Content -LiteralPath (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'MainWindow.xaml') -Raw)
    $script:Ns = New-Object System.Xml.XmlNamespaceManager($script:Xaml.NameTable)
    $script:Ns.AddNamespace('w', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
    $script:Ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
    function Get-StyleValue {
        param([string]$Key, [string]$Property)
        $style = $script:Xaml.SelectSingleNode("//w:Style[@x:Key='$Key']", $script:Ns)
        $style | Should -Not -BeNullOrEmpty -Because $Key
        $setter = $style.SelectSingleNode("w:Setter[@Property='$Property']", $script:Ns)
        if ($setter) { $setter.GetAttribute('Value') } else { $null }
    }
}

Describe 'Compact density at 1080p' {
    It 'opens at the design size of a maximized 1080p work area and keeps a small minimum' {
        $w = $script:Xaml.DocumentElement
        [int]$w.GetAttribute('Width') | Should -Be 1920
        [int]$w.GetAttribute('Height') | Should -Be 1032
        [int]$w.GetAttribute('MinWidth') | Should -BeLessOrEqual 1024
        [int]$w.GetAttribute('MinHeight') | Should -BeLessOrEqual 640
    }

    It 'sizes the sidebar view buttons between 34 and 36 px' {
        [double](Get-StyleValue 'WorkflowButtonStyle' 'Height') | Should -BeIn 34, 35, 36
    }

    It 'sizes the inline buttons and the single-line boxes between 26 and 28 px' {
        [double](Get-StyleValue 'InlineButtonStyle' 'Height') | Should -BeGreaterOrEqual 26
        [double](Get-StyleValue 'InlineButtonStyle' 'Height') | Should -BeLessOrEqual 28
        foreach ($tb in @($script:Xaml.SelectNodes('//w:TextBox[@Height]', $script:Ns))) {
            [double]$tb.GetAttribute('Height') | Should -BeIn 26, 27, 28 -Because $tb.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
        }
        foreach ($cb in @($script:Xaml.SelectNodes('//w:ComboBox[@Height]', $script:Ns))) {
            [double]$cb.GetAttribute('Height') | Should -BeIn 26, 27, 28 -Because $cb.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
        }
    }

    It 'gives every grid of the main window rows of 24 to 26 px and a 26 px column header from the one shared style' {
        [double](Get-StyleValue 'ViewDataGrid' 'RowHeight') | Should -BeGreaterOrEqual 24
        [double](Get-StyleValue 'ViewDataGrid' 'RowHeight') | Should -BeLessOrEqual 26
        [double](Get-StyleValue 'ViewDataGrid' 'ColumnHeaderHeight') | Should -Be 26
        foreach ($g in @($script:Xaml.SelectNodes('//w:DataGrid', $script:Ns))) {
            $g.GetAttribute('RowHeight') | Should -BeNullOrEmpty -Because 'the row height has one place: ViewDataGrid'
            $g.GetAttribute('ColumnHeaderHeight') | Should -BeNullOrEmpty -Because 'the header height has one place: ViewDataGrid'
            $g.GetAttribute('Style') | Should -Match 'ViewDataGrid' -Because $g.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
        }
    }

    It 'lets the two grids of the ring view share the free height with star rows (no fixed height)' {
        foreach ($name in 'dgRingPreview', 'dgRingRuns') {
            $g = $script:Xaml.SelectSingleNode("//w:DataGrid[@x:Name='$name']", $script:Ns)
            $g | Should -Not -BeNullOrEmpty -Because $name
            $g.GetAttribute('Height') | Should -BeNullOrEmpty -Because ($name + ' takes its height from a star row')
            $row = [int]$g.GetAttribute('Grid.Row')
            $defs = @($g.ParentNode.SelectNodes('w:Grid.RowDefinitions/w:RowDefinition', $script:Ns))
            $defs[$row].GetAttribute('Height') | Should -Match '\*$' -Because $name
        }
    }

    It 'gives the view its own height and the log pane the free height, so that no view leaves vacant space' {
        $rows = @($script:Xaml.SelectNodes("//w:Grid[w:Grid.RowDefinitions/w:RowDefinition[@MinHeight]]/w:Grid.RowDefinitions/w:RowDefinition", $script:Ns))
        $rows.Count | Should -BeGreaterOrEqual 5
        $rows[1].GetAttribute('Height') | Should -Be 'Auto' -Because 'the view row takes its own height'
        $rows[3].GetAttribute('Height') | Should -Match '\*$' -Because 'the log row takes the free height and has no fixed height'
        [double]$rows[3].GetAttribute('MinHeight') | Should -BeGreaterOrEqual 100
    }
}
