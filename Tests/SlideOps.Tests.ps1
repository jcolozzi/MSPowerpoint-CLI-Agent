BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'SlideOps' {
    Context 'Get-PowerPointSlide' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointSlide).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointSlide).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointSlide).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'New-PowerPointSlide' {
        It 'Has CmdletBinding' {
            (Get-Command New-PowerPointSlide).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command New-PowerPointSlide).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command New-PowerPointSlide).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Remove-PowerPointSlide' {
        It 'Has CmdletBinding' {
            (Get-Command Remove-PowerPointSlide).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Remove-PowerPointSlide).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Remove-PowerPointSlide).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Remove-PowerPointSlide).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Copy-PowerPointSlide' {
        It 'Has CmdletBinding' {
            (Get-Command Copy-PowerPointSlide).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Copy-PowerPointSlide).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Copy-PowerPointSlide).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Move-PowerPointSlide' {
        It 'Has CmdletBinding' {
            (Get-Command Move-PowerPointSlide).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Move-PowerPointSlide).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Move-PowerPointSlide).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointSlideLayout' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointSlideLayout).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointSlideLayout).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointSlideLayout).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointSlideLayout).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointSlideNotes' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointSlideNotes).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointSlideNotes).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointSlideNotes).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointSlideNotes' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointSlideNotes).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointSlideNotes).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointSlideNotes).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointSlideNotes).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Set-PowerPointSlideBackground' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointSlideBackground).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointSlideBackground).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointSlideBackground).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointSlideBackground).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointSlidePlaceholders' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointSlidePlaceholders).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointSlidePlaceholders).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointSlidePlaceholders).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }
}
