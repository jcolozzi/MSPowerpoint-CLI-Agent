BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'ExportOps' {
    Context 'Export-PowerPointToPdf' {
        It 'Has CmdletBinding' {
            (Get-Command Export-PowerPointToPdf).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Export-PowerPointToPdf).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Export-PowerPointToPdf).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Export-PowerPointToImages' {
        It 'Has CmdletBinding' {
            (Get-Command Export-PowerPointToImages).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Export-PowerPointToImages).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Export-PowerPointToImages).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Export-PowerPointToVideo' {
        It 'Has CmdletBinding' {
            (Get-Command Export-PowerPointToVideo).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Export-PowerPointToVideo).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Export-PowerPointToVideo).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Export-PowerPointToHtml' {
        It 'Has CmdletBinding' {
            (Get-Command Export-PowerPointToHtml).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Export-PowerPointToHtml).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Export-PowerPointToHtml).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Export-PowerPointSlide' {
        It 'Has CmdletBinding' {
            (Get-Command Export-PowerPointSlide).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Export-PowerPointSlide).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Export-PowerPointSlide).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Convert-PowerPointFormat' {
        It 'Has CmdletBinding' {
            (Get-Command Convert-PowerPointFormat).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Convert-PowerPointFormat).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Convert-PowerPointFormat).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Convert-PowerPointFormat).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }
}
