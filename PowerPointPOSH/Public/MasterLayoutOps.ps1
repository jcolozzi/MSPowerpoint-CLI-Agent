# Public/MasterLayoutOps.ps1 — Slide master, layout, and placeholder operations

function Get-PowerPointSlideMaster {
    <#
    .SYNOPSIS
        Get all slide masters (designs) in the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointSlideMaster -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointSlideMaster'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $masters = @()
    $idx = 0
    foreach ($design in $pres.Designs) {
        $idx++
        $layoutCount = 0
        try { $layoutCount = $design.SlideMaster.CustomLayouts.Count } catch {}

        $masterName = $null
        try { $masterName = $design.SlideMaster.Name } catch {}

        $masters += [ordered]@{
            index           = $idx
            name            = $design.Name
            slideMasterName = $masterName
            layoutCount     = $layoutCount
        }
    }

    $result = [ordered]@{
        status      = 'ok'
        masterCount = $masters.Count
        masters     = $masters
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Get-PowerPointSlideLayout {
    <#
    .SYNOPSIS
        Get all custom layouts of a slide master.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER DesignIndex
        1-based index of the design (slide master). Default: 1.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointSlideLayout -AsJson
    .EXAMPLE
        Get-PowerPointSlideLayout -DesignIndex 2 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [int]$DesignIndex = 1,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointSlideLayout'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $master  = $pres.Designs.Item($DesignIndex).SlideMaster
    $layouts = @()
    $idx = 0
    foreach ($layout in $master.CustomLayouts) {
        $idx++
        $layouts += [ordered]@{
            index  = $idx
            name   = $layout.Name
            width  = $layout.Width
            height = $layout.Height
        }
    }

    $result = [ordered]@{
        status      = 'ok'
        designIndex = $DesignIndex
        layoutCount = $layouts.Count
        layouts     = $layouts
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Set-PowerPointSlideMaster {
    <#
    .SYNOPSIS
        Apply a slide master (design) to a specific slide.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the target slide.
    .PARAMETER DesignIndex
        1-based index of the design to apply.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointSlideMaster -SlideIndex 1 -DesignIndex 2 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][int]$DesignIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Set-PowerPointSlideMaster'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $slide  = $pres.Slides.Item($SlideIndex)
    $design = $pres.Designs.Item($DesignIndex)

    if ($PSCmdlet.ShouldProcess("Slide $SlideIndex", "Apply design '$($design.Name)'")) {
        $slide.Design = $design
    }

    $result = [ordered]@{
        status = 'ok'
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function New-PowerPointCustomLayout {
    <#
    .SYNOPSIS
        Create a new custom layout on a slide master.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER DesignIndex
        1-based index of the design whose slide master gets the new layout. Default: 1.
    .PARAMETER Name
        Name for the new custom layout.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-PowerPointCustomLayout -Name 'My Layout' -AsJson
    .EXAMPLE
        New-PowerPointCustomLayout -DesignIndex 2 -Name 'Custom Wide' -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [int]$DesignIndex = 1,
        [Parameter(Mandatory)][string]$Name,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'New-PowerPointCustomLayout'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $master = $pres.Designs.Item($DesignIndex).SlideMaster
    $layoutIndex = $null

    if ($PSCmdlet.ShouldProcess("Design $DesignIndex", "Add custom layout '$Name'")) {
        $layout = $master.CustomLayouts.Add($master.CustomLayouts.Count + 1)
        $layout.Name = $Name
        $layoutIndex = $master.CustomLayouts.Count
    }

    $result = [ordered]@{
        status      = 'ok'
        layoutIndex = $layoutIndex
        name        = $Name
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Get-PowerPointPlaceholder {
    <#
    .SYNOPSIS
        Get placeholders on a slide.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointPlaceholder -SlideIndex 1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointPlaceholder'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $slide = $pres.Slides.Item($SlideIndex)
    $placeholders = @()

    foreach ($ph in $slide.Shapes.Placeholders) {
        $hasText = $false
        $text    = $null
        try {
            if ($ph.HasTextFrame -eq -1 -and $ph.TextFrame.HasText -eq -1) {
                $hasText = $true
                $rawText = $ph.TextFrame.TextRange.Text
                $text = if ($rawText.Length -gt 200) { $rawText.Substring(0, 200) + '...' } else { $rawText }
            }
        } catch {}

        $placeholders += [ordered]@{
            index   = $ph.PlaceholderFormat.Type
            name    = $ph.Name
            type    = $ph.PlaceholderFormat.Type
            left    = $ph.Left
            top     = $ph.Top
            width   = $ph.Width
            height  = $ph.Height
            hasText = $hasText
            text    = $text
        }
    }

    $result = [ordered]@{
        status           = 'ok'
        slideIndex       = $SlideIndex
        placeholderCount = $placeholders.Count
        placeholders     = $placeholders
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
