BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'AnimationOps' {
    Context 'Add-PowerPointAnimation' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointAnimation).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointAnimation).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointAnimation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointAnimation).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointAnimation' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointAnimation).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointAnimation).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointAnimation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Remove-PowerPointAnimation' {
        It 'Has CmdletBinding' {
            (Get-Command Remove-PowerPointAnimation).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Remove-PowerPointAnimation).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Remove-PowerPointAnimation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Remove-PowerPointAnimation).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Set-PowerPointAnimationTiming' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointAnimationTiming).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointAnimationTiming).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointAnimationTiming).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointAnimationTiming).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Set-PowerPointAnimationOrder' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointAnimationOrder).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointAnimationOrder).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointAnimationOrder).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointAnimationOrder).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Clear-PowerPointAnimations' {
        It 'Has CmdletBinding' {
            (Get-Command Clear-PowerPointAnimations).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Clear-PowerPointAnimations).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Clear-PowerPointAnimations).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Clear-PowerPointAnimations).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Copy-PowerPointAnimation' {
        It 'Has CmdletBinding' {
            (Get-Command Copy-PowerPointAnimation).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Copy-PowerPointAnimation).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Copy-PowerPointAnimation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Copy-PowerPointAnimation).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }
}
