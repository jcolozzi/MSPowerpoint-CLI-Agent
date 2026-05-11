BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
}

Describe 'PowerPointPOSH Module' {
    Context 'Module loads' {
        BeforeAll {
            Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
            Import-Module $modulePath -Force -ErrorAction Stop
        }
        AfterAll {
            Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
        }

        It 'Module is loaded' {
            Get-Module PowerPointPOSH | Should -Not -BeNullOrEmpty
        }

        It 'Exports expected number of public functions' {
            $exported = (Get-Module PowerPointPOSH).ExportedFunctions.Keys
            $exported.Count | Should -BeGreaterOrEqual 120
        }

        It 'Does not export cmdlets, variables, or aliases' {
            $m = Get-Module PowerPointPOSH
            $m.ExportedCmdlets.Count   | Should -Be 0
            $m.ExportedVariables.Count | Should -Be 0
            $m.ExportedAliases.Count   | Should -Be 0
        }
    }

    Context 'File structure' {
        It 'Has Private folder' {
            Test-Path (Join-Path $PSScriptRoot '..\PowerPointPOSH\Private') | Should -BeTrue
        }
        It 'Has Public folder' {
            Test-Path (Join-Path $PSScriptRoot '..\PowerPointPOSH\Public') | Should -BeTrue
        }
        It 'Has Session.ps1' {
            Test-Path (Join-Path $PSScriptRoot '..\PowerPointPOSH\Private\Session.ps1') | Should -BeTrue
        }
        It 'Has Utilities.ps1' {
            Test-Path (Join-Path $PSScriptRoot '..\PowerPointPOSH\Private\Utilities.ps1') | Should -BeTrue
        }
    }
}
