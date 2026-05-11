BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'SectionOps' {
    Context 'Get-PowerPointSection' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointSection).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointSection).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointSection).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'New-PowerPointSection' {
        It 'Has CmdletBinding' {
            (Get-Command New-PowerPointSection).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command New-PowerPointSection).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command New-PowerPointSection).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command New-PowerPointSection).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Remove-PowerPointSection' {
        It 'Has CmdletBinding' {
            (Get-Command Remove-PowerPointSection).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Remove-PowerPointSection).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Remove-PowerPointSection).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Remove-PowerPointSection).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Rename-PowerPointSection' {
        It 'Has CmdletBinding' {
            (Get-Command Rename-PowerPointSection).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Rename-PowerPointSection).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Rename-PowerPointSection).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Rename-PowerPointSection).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Move-PowerPointSection' {
        It 'Has CmdletBinding' {
            (Get-Command Move-PowerPointSection).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Move-PowerPointSection).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Move-PowerPointSection).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Move-PowerPointSection).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }
}
