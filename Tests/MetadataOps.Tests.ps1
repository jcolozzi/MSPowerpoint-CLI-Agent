BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'MetadataOps' {
    Context 'Get-PowerPointDocumentProperty' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointDocumentProperty).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointDocumentProperty).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointDocumentProperty).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointDocumentProperty' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointDocumentProperty).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointDocumentProperty).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointDocumentProperty).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointDocumentProperty).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointComment' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointComment).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointComment).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointComment).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Add-PowerPointComment' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointComment).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointComment).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointComment).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointComment).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Remove-PowerPointComment' {
        It 'Has CmdletBinding' {
            (Get-Command Remove-PowerPointComment).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Remove-PowerPointComment).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Remove-PowerPointComment).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Remove-PowerPointComment).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointTag' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointTag).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointTag).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointTag).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointTag' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointTag).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointTag).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointTag).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointTag).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }
}
