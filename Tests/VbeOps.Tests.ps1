BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'VbeOps' {
    Context 'Get-PowerPointVbaCode' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointVbaCode).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointVbaCode).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointVbaCode).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointVbaCode' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointVbaCode).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointVbaCode).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointVbaCode).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointVbaCode).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Add-PowerPointVbaModule' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointVbaModule).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointVbaModule).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointVbaModule).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointVbaModule).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Remove-PowerPointVbaModule' {
        It 'Has CmdletBinding' {
            (Get-Command Remove-PowerPointVbaModule).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Remove-PowerPointVbaModule).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Remove-PowerPointVbaModule).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Remove-PowerPointVbaModule).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointVbaModuleList' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointVbaModuleList).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointVbaModuleList).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointVbaModuleList).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Find-PowerPointVbaCode' {
        It 'Has CmdletBinding' {
            (Get-Command Find-PowerPointVbaCode).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Find-PowerPointVbaCode).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Find-PowerPointVbaCode).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Invoke-PowerPointVbaMacro' {
        It 'Has CmdletBinding' {
            (Get-Command Invoke-PowerPointVbaMacro).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Invoke-PowerPointVbaMacro).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Invoke-PowerPointVbaMacro).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Invoke-PowerPointVbaMacro).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }
}
