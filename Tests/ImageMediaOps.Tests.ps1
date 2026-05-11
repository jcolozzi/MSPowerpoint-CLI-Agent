BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'ImageMediaOps' {
    Context 'Add-PowerPointImage' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointImage).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointImage).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointImage).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointImage).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Add-PowerPointAudio' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointAudio).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointAudio).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointAudio).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointAudio).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Add-PowerPointVideo' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointVideo).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointVideo).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointVideo).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointVideo).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointMedia' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointMedia).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointMedia).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointMedia).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointMediaProperties' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointMediaProperties).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointMediaProperties).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointMediaProperties).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointMediaProperties).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Remove-PowerPointMedia' {
        It 'Has CmdletBinding' {
            (Get-Command Remove-PowerPointMedia).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Remove-PowerPointMedia).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Remove-PowerPointMedia).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Remove-PowerPointMedia).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Export-PowerPointSlideImage' {
        It 'Has CmdletBinding' {
            (Get-Command Export-PowerPointSlideImage).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Export-PowerPointSlideImage).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Export-PowerPointSlideImage).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }
}
