BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'PresentationOps' {
    Context 'Open-PowerPointPresentation' {
        It 'Has CmdletBinding' {
            (Get-Command Open-PowerPointPresentation).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Open-PowerPointPresentation).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Open-PowerPointPresentation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'New-PowerPointPresentation' {
        It 'Has CmdletBinding' {
            (Get-Command New-PowerPointPresentation).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command New-PowerPointPresentation).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command New-PowerPointPresentation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Save-PowerPointPresentation' {
        It 'Has CmdletBinding' {
            (Get-Command Save-PowerPointPresentation).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Save-PowerPointPresentation).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Save-PowerPointPresentation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Close-PowerPointPresentation' {
        It 'Has CmdletBinding' {
            (Get-Command Close-PowerPointPresentation).CmdletBinding | Should -BeTrue
        }
        It 'Does not have PresentationPath parameter' {
            (Get-Command Close-PowerPointPresentation).Parameters['PresentationPath'] | Should -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Close-PowerPointPresentation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Get-PowerPointPresentationInfo' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointPresentationInfo).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointPresentationInfo).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointPresentationInfo).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Copy-PowerPointPresentation' {
        It 'Has CmdletBinding' {
            (Get-Command Copy-PowerPointPresentation).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Copy-PowerPointPresentation).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Copy-PowerPointPresentation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Convert-PowerPointPresentation' {
        It 'Has CmdletBinding' {
            (Get-Command Convert-PowerPointPresentation).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Convert-PowerPointPresentation).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Convert-PowerPointPresentation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Repair-PowerPointPresentation' {
        It 'Has CmdletBinding' {
            (Get-Command Repair-PowerPointPresentation).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Repair-PowerPointPresentation).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Repair-PowerPointPresentation).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }
}
