BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'MasterLayoutOps' {
    Context 'Get-PowerPointSlideMaster' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointSlideMaster).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointSlideMaster).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointSlideMaster).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Get-PowerPointSlideLayout' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointSlideLayout).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointSlideLayout).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointSlideLayout).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointSlideMaster' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointSlideMaster).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointSlideMaster).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointSlideMaster).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointSlideMaster).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'New-PowerPointCustomLayout' {
        It 'Has CmdletBinding' {
            (Get-Command New-PowerPointCustomLayout).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command New-PowerPointCustomLayout).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command New-PowerPointCustomLayout).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command New-PowerPointCustomLayout).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointPlaceholder' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointPlaceholder).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointPlaceholder).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointPlaceholder).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }
}
