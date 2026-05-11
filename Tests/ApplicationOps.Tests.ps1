BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'ApplicationOps' {
    Context 'Get-PowerPointApplicationInfo' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointApplicationInfo).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointApplicationInfo).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointApplicationInfo).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointOption' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointOption).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointOption).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointOption).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointOption).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointTip' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointTip).CmdletBinding | Should -BeTrue
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointTip).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Does not have PresentationPath parameter' {
            (Get-Command Get-PowerPointTip).Parameters.ContainsKey('PresentationPath') | Should -BeFalse
        }
    }
}
