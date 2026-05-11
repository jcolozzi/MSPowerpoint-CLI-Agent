# Public/SmartArtOps.ps1 — SmartArt manipulation operations

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointSmartArt
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointSmartArt {
    <#
    .SYNOPSIS
        Add a SmartArt graphic to a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER LayoutName
        Name of the SmartArt layout (e.g. 'Basic Block List', 'Organization Chart').
    .PARAMETER Left
        Left position in points (default 100).
    .PARAMETER Top
        Top position in points (default 100).
    .PARAMETER Width
        Width in points (default 400).
    .PARAMETER Height
        Height in points (default 300).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointSmartArt -SlideIndex 1 -LayoutName "Basic Block List" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$LayoutName,
        [double]$Left   = 100,
        [double]$Top    = 100,
        [double]$Width  = 400,
        [double]$Height = 300,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $layout  = $null
    $layouts = $app.SmartArtLayouts
    foreach ($l in $layouts) {
        if ($l.Name -eq $LayoutName) { $layout = $l; break }
    }
    if (-not $layout) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentException]::new("SmartArt layout '$LayoutName' not found."),
            'LayoutNotFound',
            [System.Management.Automation.ErrorCategory]::ObjectNotFound, $LayoutName)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex", "Add SmartArt '$LayoutName'")) { return }

    $shape = $slide.Shapes.AddSmartArt($layout, $Left, $Top, $Width, $Height)

    $result = [ordered]@{
        status = 'added'
        name   = $shape.Name
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointSmartArt
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointSmartArt {
    <#
    .SYNOPSIS
        Get SmartArt information from a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the SmartArt shape.
    .PARAMETER ShapeIndex
        1-based index of the SmartArt shape.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointSmartArt -SlideIndex 1 -ShapeName "SmartArt 1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasSmartArt -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain SmartArt."),
            'NoSmartArt',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $smartArt  = $shape.SmartArt
    $nodeCount = $smartArt.AllNodes.Count
    $nodes     = @()
    for ($i = 1; $i -le $nodeCount; $i++) {
        $node = $smartArt.AllNodes.Item($i)
        $nodes += [ordered]@{
            index = $i
            text  = $node.TextFrame2.TextRange.Text
        }
    }

    $result = [ordered]@{
        shapeName  = $shape.Name
        layoutName = $smartArt.Layout.Name
        nodeCount  = $nodeCount
        nodes      = $nodes
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointSmartArtLayout
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointSmartArtLayout {
    <#
    .SYNOPSIS
        Change the SmartArt layout of a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the SmartArt shape.
    .PARAMETER ShapeIndex
        1-based index of the SmartArt shape.
    .PARAMETER LayoutName
        Name of the new SmartArt layout.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointSmartArtLayout -SlideIndex 1 -ShapeName "SmartArt 1" -LayoutName "Organization Chart" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][string]$LayoutName,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasSmartArt -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain SmartArt."),
            'NoSmartArt',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $layout  = $null
    $layouts = $app.SmartArtLayouts
    foreach ($l in $layouts) {
        if ($l.Name -eq $LayoutName) { $layout = $l; break }
    }
    if (-not $layout) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentException]::new("SmartArt layout '$LayoutName' not found."),
            'LayoutNotFound',
            [System.Management.Automation.ErrorCategory]::ObjectNotFound, $LayoutName)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("SmartArt '$target' on slide $SlideIndex", "Change layout to '$LayoutName'")) { return }

    $shape.SmartArt.Layout = $layout

    $result = [ordered]@{
        status     = 'updated'
        shape      = $shape.Name
        layoutName = $LayoutName
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointSmartArtNode
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointSmartArtNode {
    <#
    .SYNOPSIS
        Set the text of a SmartArt node.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the SmartArt shape.
    .PARAMETER ShapeIndex
        1-based index of the SmartArt shape.
    .PARAMETER NodeIndex
        1-based index of the node in AllNodes collection.
    .PARAMETER Text
        Text to set on the node.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointSmartArtNode -SlideIndex 1 -ShapeName "SmartArt 1" -NodeIndex 1 -Text "CEO" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][int]$NodeIndex,
        [Parameter(Mandatory)][string]$Text,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasSmartArt -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain SmartArt."),
            'NoSmartArt',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Node $NodeIndex in SmartArt '$target' on slide $SlideIndex", "Set text")) { return }

    $shape.SmartArt.AllNodes.Item($NodeIndex).TextFrame2.TextRange.Text = $Text

    $result = [ordered]@{
        status    = 'updated'
        shape     = $shape.Name
        nodeIndex = $NodeIndex
        text      = $Text
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
