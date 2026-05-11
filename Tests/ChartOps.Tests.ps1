BeforeAll {
    $modulePath = Join-Path $PSScriptRoot '..\PowerPointPOSH\PowerPointPOSH.psd1'
    $modulePath = (Resolve-Path $modulePath).Path
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
    Import-Module $modulePath -Force -ErrorAction Stop
}
AfterAll {
    Get-Module PowerPointPOSH -ErrorAction SilentlyContinue | Remove-Module -Force
}

Describe 'ChartOps' {
    Context 'Add-PowerPointChart' {
        It 'Has CmdletBinding' {
            (Get-Command Add-PowerPointChart).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Add-PowerPointChart).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Add-PowerPointChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Add-PowerPointChart).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Get-PowerPointChart' {
        It 'Has CmdletBinding' {
            (Get-Command Get-PowerPointChart).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Get-PowerPointChart).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Get-PowerPointChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Set-PowerPointChart' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointChart).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointChart).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointChart).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Set-PowerPointChartData' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointChartData).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointChartData).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointChartData).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointChartData).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Set-PowerPointChartSeries' {
        It 'Has CmdletBinding' {
            (Get-Command Set-PowerPointChartSeries).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Set-PowerPointChartSeries).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Set-PowerPointChartSeries).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
        It 'Has SupportsShouldProcess' {
            (Get-Command Set-PowerPointChartSeries).Parameters.ContainsKey('WhatIf') | Should -BeTrue
        }
    }

    Context 'Export-PowerPointChart' {
        It 'Has CmdletBinding' {
            (Get-Command Export-PowerPointChart).CmdletBinding | Should -BeTrue
        }
        It 'Has PresentationPath parameter' {
            (Get-Command Export-PowerPointChart).Parameters['PresentationPath'] | Should -Not -BeNullOrEmpty
        }
        It 'Has AsJson parameter' {
            (Get-Command Export-PowerPointChart).Parameters['AsJson'] | Should -Not -BeNullOrEmpty
        }
    }
}
