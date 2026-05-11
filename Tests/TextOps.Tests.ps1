BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'TextOps' {
    Context 'Get-PowerPointText' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointText).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointText).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointText).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointText' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointText).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointText).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointText).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointText).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Format-PowerPointTextFont' {
        It 'Has CmdletBinding' {
            (Get-Command Format-PowerPointTextFont).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Format-PowerPointTextFont).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Format-PowerPointTextFont).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Format-PowerPointTextFont).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Format-PowerPointTextParagraph' {
        It 'Has CmdletBinding' {
            (Get-Command Format-PowerPointTextParagraph).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Format-PowerPointTextParagraph).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Format-PowerPointTextParagraph).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Format-PowerPointTextParagraph).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Add-PowerPointBullet' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointBullet).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointBullet).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointBullet).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointBullet).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Find-PowerPointText' {
        It 'Has CmdletBinding' {
            (Get-Command Find-PowerPointText).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Find-PowerPointText).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Find-PowerPointText).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointTextReplace' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointTextReplace).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointTextReplace).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointTextReplace).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointTextReplace).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointTextAll' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointTextAll).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointTextAll).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointTextAll).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }
}
