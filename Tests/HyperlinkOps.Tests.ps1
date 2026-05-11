BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'HyperlinkOps' {
    Context 'Add-PowerPointHyperlink' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointHyperlink).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointHyperlink).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointHyperlink).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointHyperlink).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointHyperlink' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointHyperlink).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointHyperlink).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointHyperlink).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Remove-PowerPointHyperlink' {
        It 'Has CmdletBinding' {
            (Get-Command Remove-PowerPointHyperlink).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Remove-PowerPointHyperlink).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Remove-PowerPointHyperlink).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Remove-PowerPointHyperlink).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }
}
