# Public/PresentationOps.ps1 — Presentation lifecycle: open, new, save, close, info, copy, convert, repair

function Open-PowerPointPresentation {
    <#
    .SYNOPSIS
        Open a PowerPoint presentation and establish a COM session.
    .PARAMETER PresentationPath
        Path to the .pptx/.pptm/.ppt file.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Open-PowerPointPresentation -PresentationPath "C:\deck.pptx" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$PresentationPath,
        [switch]$AsJson
    )

    $PresentationPath = [System.IO.Path]::GetFullPath($PresentationPath)
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $activeSlide = $null
    try { $activeSlide = $app.ActiveWindow.View.Slide.SlideIndex } catch {}

    $result = [ordered]@{
        status      = 'ok'
        name        = $pres.Name
        path        = $pres.FullName
        slideCount  = $pres.Slides.Count
        slideWidth  = $pres.PageSetup.SlideWidth
        slideHeight = $pres.PageSetup.SlideHeight
        activeSlide = $activeSlide
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function New-PowerPointPresentation {
    <#
    .SYNOPSIS
        Create a new blank PowerPoint presentation at the specified path.
    .PARAMETER PresentationPath
        Full path for the new presentation file.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-PowerPointPresentation -PresentationPath "C:\decks\new.pptx"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$PresentationPath,
        [switch]$AsJson
    )

    $resolved = [System.IO.Path]::GetFullPath($PresentationPath)
    if (Test-Path -LiteralPath $resolved) {
        throw "File already exists: $resolved"
    }

    $dir = [System.IO.Path]::GetDirectoryName($resolved)
    if (-not (Test-Path -LiteralPath $dir)) {
        $null = New-Item -Path $dir -ItemType Directory -Force
    }

    # Launch PowerPoint if needed
    if ($null -eq $script:PowerPointSession.App) {
        try {
            $script:PowerPointSession.App = New-Object -ComObject 'PowerPoint.Application'
        } catch {
            throw "Failed to create PowerPoint.Application COM object. Is Microsoft PowerPoint installed? Error: $_"
        }
        $script:PowerPointSession.App.DisplayAlerts = 0
        Set-PowerPointVisibleBestEffort -Visible $true
    }

    $app  = $script:PowerPointSession.App
    $pres = $app.Presentations.Add()

    # Determine file format from extension
    $ext = [System.IO.Path]::GetExtension($resolved).ToLower().TrimStart('.')
    $fmt = if ($script:PPT_FILE_FORMAT.ContainsKey($ext)) {
        $script:PPT_FILE_FORMAT[$ext]
    } else {
        $script:PPT_FILE_FORMAT['pptx']
    }

    $pres.SaveAs($resolved, $fmt)
    $script:PowerPointSession.PresentationPath = $resolved

    $result = [ordered]@{
        status     = 'ok'
        path       = $pres.FullName
        slideCount = $pres.Slides.Count
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Save-PowerPointPresentation {
    <#
    .SYNOPSIS
        Save the current presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Save-PowerPointPresentation -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Save-PowerPointPresentation'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $pres.Save()

    $lastSaved = $null
    try { $lastSaved = $pres.BuiltInDocumentProperties('Last Save Time').Value.ToString('o') } catch {}

    $result = [ordered]@{
        status    = 'ok'
        path      = $pres.FullName
        lastSaved = $lastSaved
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Close-PowerPointPresentation {
    <#
    .SYNOPSIS
        Close the current presentation and optionally quit PowerPoint.
    .PARAMETER Save
        Save the presentation before closing.
    .PARAMETER QuitApp
        Quit PowerPoint and release the COM object after closing.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Close-PowerPointPresentation -Save -QuitApp
    .EXAMPLE
        Close-PowerPointPresentation
    #>
    [CmdletBinding()]
    param(
        [switch]$Save,
        [switch]$QuitApp,
        [switch]$AsJson
    )

    if ($null -ne $script:PowerPointSession.App) {
        # Close the active presentation
        try {
            $pres = $script:PowerPointSession.App.ActivePresentation
            if ($Save) { $pres.Save() }
            $pres.Close()
        } catch {
            Write-Verbose "Error closing presentation: $_"
        }

        $script:PowerPointSession.PresentationPath = $null

        if ($QuitApp) {
            try { $script:PowerPointSession.App.DisplayAlerts = 0 } catch {}
            try { $script:PowerPointSession.App.Quit() } catch {}
            Start-Sleep -Milliseconds 500
            try {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:PowerPointSession.App)
            } catch {}
            $script:PowerPointSession.App = $null
        }
    }

    $result = [ordered]@{
        status  = if ($QuitApp) { 'closed_and_quit' } else { 'closed' }
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Get-PowerPointPresentationInfo {
    <#
    .SYNOPSIS
        Return detailed information about the current presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointPresentationInfo -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointPresentationInfo'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $hasVBA = $false
    try {
        $null = $pres.VBProject.VBComponents.Count
        $hasVBA = $true
    } catch {
        $hasVBA = $false
    }

    $defaultShapeInfo = $null
    try {
        $ds = $pres.DefaultShape
        $defaultShapeInfo = [ordered]@{
            fontName = $ds.TextFrame.TextRange.Font.Name
            fontSize = $ds.TextFrame.TextRange.Font.Size
            fontBold = [bool]$ds.TextFrame.TextRange.Font.Bold
        }
    } catch {}

    $result = [ordered]@{
        status       = 'ok'
        name         = $pres.Name
        fullPath     = $pres.FullName
        slideCount   = $pres.Slides.Count
        slideWidth   = $pres.PageSetup.SlideWidth
        slideHeight  = $pres.PageSetup.SlideHeight
        hasVBA       = $hasVBA
        readOnly     = [bool]$pres.ReadOnly
        saved        = [bool]$pres.Saved
        defaultShape = $defaultShapeInfo
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Copy-PowerPointPresentation {
    <#
    .SYNOPSIS
        Save a copy of the current presentation to a new path (SaveCopyAs).
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER DestinationPath
        Full path for the copy.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Copy-PowerPointPresentation -DestinationPath "C:\backup\deck_copy.pptx" -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$DestinationPath,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Copy-PowerPointPresentation'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $resolvedDest = [System.IO.Path]::GetFullPath($DestinationPath)
    $destDir = [System.IO.Path]::GetDirectoryName($resolvedDest)
    if (-not (Test-Path -LiteralPath $destDir)) {
        $null = New-Item -Path $destDir -ItemType Directory -Force
    }

    $pres.SaveCopyAs($resolvedDest)

    $result = [ordered]@{
        status      = 'ok'
        source      = $pres.FullName
        destination = $resolvedDest
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Convert-PowerPointPresentation {
    <#
    .SYNOPSIS
        Save the presentation in a different format (SaveAs with format).
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER DestinationPath
        Full path for the converted file.
    .PARAMETER Format
        Target format key (e.g., 'pdf', 'pptx', 'png', 'mp4'). Maps to PPT_FILE_FORMAT.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Convert-PowerPointPresentation -DestinationPath "C:\out\deck.pdf" -Format 'pdf' -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$DestinationPath,
        [Parameter(Mandatory)][string]$Format,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Convert-PowerPointPresentation'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $fmtValue = Resolve-EnumValue -Map $script:PPT_FILE_FORMAT -Key $Format

    $resolvedDest = [System.IO.Path]::GetFullPath($DestinationPath)
    $destDir = [System.IO.Path]::GetDirectoryName($resolvedDest)
    if (-not (Test-Path -LiteralPath $destDir)) {
        $null = New-Item -Path $destDir -ItemType Directory -Force
    }

    $pres.SaveAs($resolvedDest, $fmtValue)

    # Reopen the original since SaveAs changes the active presentation
    $script:PowerPointSession.PresentationPath = $null
    $null = Connect-PowerPointPresentation -PresentationPath $PresentationPath

    $result = [ordered]@{
        status      = 'ok'
        source      = $PresentationPath
        destination = $resolvedDest
        format      = $Format
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Repair-PowerPointPresentation {
    <#
    .SYNOPSIS
        Open, force-save, and close a presentation to attempt repair.
    .PARAMETER PresentationPath
        Path to the presentation to repair.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Repair-PowerPointPresentation -PresentationPath "C:\corrupt.pptx" -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$PresentationPath,
        [switch]$AsJson
    )

    $PresentationPath = [System.IO.Path]::GetFullPath($PresentationPath)
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $pres.Save()
    $pres.Close()

    $script:PowerPointSession.PresentationPath = $null

    $result = [ordered]@{
        status = 'repaired'
        path   = $PresentationPath
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
