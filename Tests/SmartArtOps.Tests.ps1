BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'SmartArtOps' {
    Context 'Add-PowerPointSmartArt' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointSmartArt).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointSmartArt).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointSmartArt).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointSmartArt).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointSmartArt' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointSmartArt).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointSmartArt).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointSmartArt).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointSmartArtLayout' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointSmartArtLayout).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointSmartArtLayout).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointSmartArtLayout).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointSmartArtLayout).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Set-PowerPointSmartArtNode' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointSmartArtNode).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointSmartArtNode).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointSmartArtNode).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointSmartArtNode).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }
}
