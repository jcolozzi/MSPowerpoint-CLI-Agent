BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'ShapeOps' {
    Context 'Get-PowerPointShape' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointShape).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointShape).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointShape).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Add-PowerPointShape' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointShape).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointShape).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointShape).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointShape).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Add-PowerPointTextBox' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointTextBox).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointTextBox).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointTextBox).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointTextBox).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Add-PowerPointLine' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointLine).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointLine).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointLine).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointLine).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Remove-PowerPointShape' {
        It 'Has CmdletBinding' {
            (Get-Command Remove-PowerPointShape).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Remove-PowerPointShape).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Remove-PowerPointShape).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Remove-PowerPointShape).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Set-PowerPointShapeProperties' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointShapeProperties).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointShapeProperties).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointShapeProperties).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointShapeProperties).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Copy-PowerPointShape' {
        It 'Has CmdletBinding' {
            (Get-Command Copy-PowerPointShape).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Copy-PowerPointShape).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Copy-PowerPointShape).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Group-PowerPointShapes' {
        It 'Has CmdletBinding' {
            (Get-Command Group-PowerPointShapes).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Group-PowerPointShapes).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Group-PowerPointShapes).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Ungroup-PowerPointShapes' {
        It 'Has CmdletBinding' {
            (Get-Command Ungroup-PowerPointShapes).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Ungroup-PowerPointShapes).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Ungroup-PowerPointShapes).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointShapeZOrder' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointShapeZOrder).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointShapeZOrder).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointShapeZOrder).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointShapeZOrder).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointShapeList' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointShapeList).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointShapeList).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointShapeList).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Find-PowerPointShape' {
        It 'Has CmdletBinding' {
            (Get-Command Find-PowerPointShape).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Find-PowerPointShape).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Find-PowerPointShape).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Align-PowerPointShapes' {
        It 'Has CmdletBinding' {
            (Get-Command Align-PowerPointShapes).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Align-PowerPointShapes).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Align-PowerPointShapes).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Align-PowerPointShapes).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Distribute-PowerPointShapes' {
        It 'Has CmdletBinding' {
            (Get-Command Distribute-PowerPointShapes).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Distribute-PowerPointShapes).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Distribute-PowerPointShapes).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Distribute-PowerPointShapes).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }
}
