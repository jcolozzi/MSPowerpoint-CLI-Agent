BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'SlideShowOps' {
    Context 'Start-PowerPointSlideShow' {
        It 'Has CmdletBinding' {
            (Get-Command Start-PowerPointSlideShow).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Start-PowerPointSlideShow).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Start-PowerPointSlideShow).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Stop-PowerPointSlideShow' {
        It 'Has CmdletBinding' {
            (Get-Command Stop-PowerPointSlideShow).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Stop-PowerPointSlideShow).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Stop-PowerPointSlideShow).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointSlideShowSettings' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointSlideShowSettings).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointSlideShowSettings).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointSlideShowSettings).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointSlideShowSettings).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointSlideShowInfo' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointSlideShowInfo).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointSlideShowInfo).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointSlideShowInfo).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Step-PowerPointSlideShow' {
        It 'Has CmdletBinding' {
            (Get-Command Step-PowerPointSlideShow).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Step-PowerPointSlideShow).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Step-PowerPointSlideShow).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointPresenterView' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointPresenterView).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointPresenterView).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointPresenterView).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointPresenterView).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }
}
