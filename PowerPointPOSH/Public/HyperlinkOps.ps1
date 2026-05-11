# Public/HyperlinkOps.ps1 — Hyperlink operations: add, get, remove

function Add-PowerPointHyperlink {
    <#
    .SYNOPSIS
        Add a hyperlink to a shape on a slide.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the target slide.
    .PARAMETER ShapeName
        Name of the shape to add the hyperlink to.
    .PARAMETER ShapeIndex
        1-based index of the shape to add the hyperlink to.
    .PARAMETER Address
        URL or file path for the hyperlink.
    .PARAMETER SubAddress
        Sub-address (e.g., slide number, named location).
    .PARAMETER ScreenTip
        Tooltip text displayed on hover.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointHyperlink -SlideIndex 1 -ShapeName 'Title 1' -Address 'https://example.com' -AsJson
    .EXAMPLE
        Add-PowerPointHyperlink -SlideIndex 2 -ShapeIndex 3 -Address 'https://example.com' -SubAddress 'slide2' -ScreenTip 'Go to example'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][string]$Address,
        [string]$SubAddress,
        [string]$ScreenTip,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Add-PowerPointHyperlink'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "Shape $ShapeIndex" }

    if ($PSCmdlet.ShouldProcess("Slide $SlideIndex / $target", "Add hyperlink to '$Address'")) {
        # ppMouseClick = 1
        $shape.ActionSettings(1).Hyperlink.Address = $Address

        if ($PSBoundParameters.ContainsKey('SubAddress')) {
            $shape.ActionSettings(1).Hyperlink.SubAddress = $SubAddress
        }
        if ($PSBoundParameters.ContainsKey('ScreenTip')) {
            $shape.ActionSettings(1).Hyperlink.ScreenTip = $ScreenTip
        }
    }

    $result = [ordered]@{
        status = 'ok'
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Get-PowerPointHyperlink {
    <#
    .SYNOPSIS
        Get hyperlinks from shapes on a slide.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide to inspect.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointHyperlink -SlideIndex 1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointHyperlink'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $slide = $pres.Slides.Item($SlideIndex)
    $links = @()

    foreach ($shape in $slide.Shapes) {
        $addr = $null
        try { $addr = $shape.ActionSettings(1).Hyperlink.Address } catch {}
        if (-not [string]::IsNullOrWhiteSpace($addr)) {
            $subAddr  = try { $shape.ActionSettings(1).Hyperlink.SubAddress } catch { $null }
            $tip      = try { $shape.ActionSettings(1).Hyperlink.ScreenTip } catch { $null }
            $links += [ordered]@{
                shapeName  = $shape.Name
                address    = $addr
                subAddress = $subAddr
                screenTip  = $tip
            }
        }
    }

    $result = [ordered]@{
        status    = 'ok'
        linkCount = $links.Count
        links     = $links
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Remove-PowerPointHyperlink {
    <#
    .SYNOPSIS
        Remove a hyperlink from a shape on a slide.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the target slide.
    .PARAMETER ShapeName
        Name of the shape to remove the hyperlink from.
    .PARAMETER ShapeIndex
        1-based index of the shape to remove the hyperlink from.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-PowerPointHyperlink -SlideIndex 1 -ShapeName 'Title 1' -AsJson
    .EXAMPLE
        Remove-PowerPointHyperlink -SlideIndex 2 -ShapeIndex 3 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Remove-PowerPointHyperlink'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "Shape $ShapeIndex" }

    if ($PSCmdlet.ShouldProcess("Slide $SlideIndex / $target", 'Remove hyperlink')) {
        # ppActionNone = 0
        $shape.ActionSettings(1).Action = 0
    }

    $result = [ordered]@{
        status = 'ok'
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
