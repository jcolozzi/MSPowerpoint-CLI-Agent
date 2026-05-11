# Private/Session.ps1 — COM session management helpers

function Get-RunningComApp {
    <#
    .SYNOPSIS
        Try to attach to an already-running COM application via the Running Object Table.
        Returns the COM object or $null.  Works on Windows PowerShell 5.1 (Desktop);
        gracefully degrades on PowerShell 7+ where [Marshal]::GetActiveObject is unavailable.
    #>
    param(
        [Parameter(Mandatory)][string]$ProgId,
        [Parameter(Mandatory)][string]$ProcessName
    )

    # Fast exit: if the host process isn't running, skip the COM probe entirely
    if (-not (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)) {
        Write-Verbose "Get-RunningComApp: no $ProcessName process found — skipping ROT lookup."
        return $null
    }

    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
        Write-Verbose "Get-RunningComApp: attached to existing $ProgId instance."
        return $app
    }
    catch [System.Management.Automation.MethodException] {
        # .NET Core / PS7 — GetActiveObject does not exist
        Write-Verbose "Get-RunningComApp: [Marshal]::GetActiveObject unavailable (PowerShell $($PSVersionTable.PSVersion)) — will create new instance."
        return $null
    }
    catch {
        # No ROT entry, or stale/dead entry
        Write-Verbose "Get-RunningComApp: could not attach to $ProgId — $($_.Exception.Message)"
        return $null
    }
}

function Test-PowerPointAlive {
    <#
    .SYNOPSIS
        Best-effort COM liveness check for PowerPoint.
    #>
    if ($null -eq $script:PowerPointSession.App) { return $false }
    $alive = $false
    try {
        $null = $script:PowerPointSession.App.HWND
        $alive = $true
    }
    catch {
        try {
            $null = $script:PowerPointSession.App.Version
            $alive = $true
        }
        catch {
            $alive = $false
        }
    }
    return $alive
}

function Set-PowerPointVisibleBestEffort {
    <#
    .SYNOPSIS
        Try to set PowerPoint visibility. Never fail startup if unsupported.
    #>
    param([bool]$Visible = $true)
    if ($null -eq $script:PowerPointSession.App) { return }
    try {
        $script:PowerPointSession.App.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
    } catch {
        try {
            $script:PowerPointSession.App.Visible = -1  # msoTrue
        } catch {
            Write-Verbose "Could not set PowerPoint.Visible=$Visible (continuing): $_"
        }
    }
}

function Connect-PowerPointPresentation {
    <#
    .SYNOPSIS
        Internal: ensure PowerPoint COM is running and the requested presentation is open.
        Returns the COM Application object.
        Tries to attach to an already-running PowerPoint instance (GetObject-first)
        before creating a new one, to prevent duplicate instances.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$PresentationPath,
        [switch]$ForceNewInstance
    )

    $resolved = [System.IO.Path]::GetFullPath($PresentationPath)

    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw "Presentation file not found: $resolved"
    }

    # If we have an existing session, check liveness
    if ($null -ne $script:PowerPointSession.App) {
        if (-not (Test-PowerPointAlive)) {
            Write-Verbose 'COM session stale — auto-reconnecting...'
            $script:PowerPointSession.App              = $null
            $script:PowerPointSession.PresentationPath = $null
            $script:PowerPointSession.OwnsApp          = $false
        }
    }

    # Acquire PowerPoint instance if needed (GetObject-first, then New-Object)
    if ($null -eq $script:PowerPointSession.App) {
        $adopted = $false

        # Try to attach to an existing PowerPoint instance via the ROT
        if (-not $ForceNewInstance) {
            $existing = Get-RunningComApp -ProgId 'PowerPoint.Application' -ProcessName 'POWERPNT'
            if ($null -ne $existing) {
                Write-Verbose 'Adopting existing PowerPoint instance'
                $script:PowerPointSession.App     = $existing
                $script:PowerPointSession.OwnsApp = $false
                $adopted = $true
                $script:PowerPointSession.App.DisplayAlerts = 0  # ppAlertsNone
                Set-PowerPointVisibleBestEffort -Visible $true
                Write-Verbose 'Adopted existing PowerPoint instance OK'
            }
        }

        # Fall back to creating a new instance
        if (-not $adopted) {
            Write-Verbose 'Launching PowerPoint.Application...'
            try {
                $script:PowerPointSession.App = New-Object -ComObject 'PowerPoint.Application'
            } catch {
                throw "Failed to create PowerPoint.Application COM object. Is Microsoft PowerPoint installed? Error: $_"
            }
            $script:PowerPointSession.OwnsApp = $true
            $script:PowerPointSession.App.DisplayAlerts = 0  # ppAlertsNone
            Set-PowerPointVisibleBestEffort -Visible $true
            Write-Verbose 'PowerPoint launched OK'
        }
    }

    # Switch presentation if needed
    if ($script:PowerPointSession.PresentationPath -ne $resolved) {
        # Check if the presentation is already open in this instance
        $alreadyOpen = $false
        try {
            foreach ($pres in $script:PowerPointSession.App.Presentations) {
                if ($pres.FullName -eq $resolved) {
                    Write-Verbose "Presentation already open in this instance: $resolved"
                    $alreadyOpen = $true
                    break
                }
            }
        } catch {
            Write-Verbose "Error checking open presentations: $_"
        }

        # Close previous presentation (only if different from the one we're opening)
        if (-not $alreadyOpen -and $null -ne $script:PowerPointSession.PresentationPath) {
            Write-Verbose "Closing previous presentation: $($script:PowerPointSession.PresentationPath)"
            try {
                foreach ($pres in $script:PowerPointSession.App.Presentations) {
                    if ($pres.FullName -eq $script:PowerPointSession.PresentationPath) {
                        $pres.Close()
                        break
                    }
                }
            } catch {
                Write-Verbose "Error closing previous presentation: $_"
            }
        }

        # Open presentation if not already open
        if (-not $alreadyOpen) {
            Write-Verbose "Opening presentation: $resolved"
            try {
                $null = $script:PowerPointSession.App.Presentations.Open($resolved)
            } catch {
                throw "Failed to open presentation '$resolved': $_"
            }
        }

        $script:PowerPointSession.PresentationPath = $resolved
        Write-Verbose "Presentation opened: $resolved"
    }

    return $script:PowerPointSession.App
}

function Resolve-SessionPresentationPath {
    <#
    .SYNOPSIS
        Resolve the presentation path from explicit parameter or session state.
    #>
    param(
        [string]$PresentationPath,
        [string]$CallerName = 'Unknown'
    )

    if (-not [string]::IsNullOrWhiteSpace($PresentationPath)) {
        return [System.IO.Path]::GetFullPath($PresentationPath)
    }

    if (-not [string]::IsNullOrWhiteSpace($script:PowerPointSession.PresentationPath)) {
        return $script:PowerPointSession.PresentationPath
    }

    throw "$CallerName requires -PresentationPath (no presentation is currently open in session)."
}
