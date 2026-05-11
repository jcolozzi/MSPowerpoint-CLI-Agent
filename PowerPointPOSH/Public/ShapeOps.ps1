# Public/ShapeOps.ps1 — Shape manipulation operations

# ── Private helpers (visible to all dot-sourced files at runtime) ──────────

function Get-SlideShape {
    param($Slide, $ShapeName, $ShapeIndex)
    if ($ShapeName)  { return $Slide.Shapes.Item($ShapeName) }
    if ($ShapeIndex) { return $Slide.Shapes.Item($ShapeIndex) }
    throw "Either -ShapeName or -ShapeIndex must be specified."
}

function ConvertFrom-RGBString {
    param([Parameter(Mandatory)][string]$RGB)
    $parts = $RGB -split ','
    if ($parts.Count -ne 3) {
        throw "Invalid RGB format '$RGB'. Expected 'R,G,B' (e.g. '255,0,0')."
    }
    return [int]$parts[0].Trim() + ([int]$parts[1].Trim() * 256) + ([int]$parts[2].Trim() * 65536)
}

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointShape
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointShape {
    <#
    .SYNOPSIS
        Get shape information from a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape to retrieve.
    .PARAMETER ShapeIndex
        1-based index of the shape to retrieve.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointShape -PresentationPath "C:\deck.pptx" -SlideIndex 1 -ShapeName "Title 1" -AsJson
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

    $getShapeInfo = {
        param($s)
        $hasText  = $false
        try { if ($s.HasTextFrame -eq -1 -and $s.TextFrame.HasText -eq -1) { $hasText = $true } } catch {}
        $hasTable = try { $s.HasTable -eq -1 } catch { $false }
        $hasChart = try { $s.HasChart -eq -1 } catch { $false }
        [ordered]@{
            name           = $s.Name
            type           = [int]$s.Type
            left           = $s.Left
            top            = $s.Top
            width          = $s.Width
            height         = $s.Height
            rotation       = $s.Rotation
            hasText        = $hasText
            hasTable       = $hasTable
            hasChart       = $hasChart
            visible        = $s.Visible -eq -1
            zOrderPosition = $s.ZOrderPosition
        }
    }

    if ($ShapeName -or $ShapeIndex) {
        $shape  = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex
        $result = & $getShapeInfo $shape
    }
    else {
        $result = @()
        foreach ($s in $slide.Shapes) { $result += & $getShapeInfo $s }
    }

    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointShape
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointShape {
    <#
    .SYNOPSIS
        Add an AutoShape to a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeType
        Shape type key (e.g. 'rectangle', 'oval'). Resolved via PPT_SHAPE_TYPE map.
    .PARAMETER Left
        Left position in points.
    .PARAMETER Top
        Top position in points.
    .PARAMETER Width
        Width in points.
    .PARAMETER Height
        Height in points.
    .PARAMETER Name
        Optional name to assign to the new shape.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointShape -SlideIndex 1 -ShapeType rectangle -Left 100 -Top 100 -Width 200 -Height 100 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$ShapeType,
        [Parameter(Mandatory)][double]$Left,
        [Parameter(Mandatory)][double]$Top,
        [Parameter(Mandatory)][double]$Width,
        [Parameter(Mandatory)][double]$Height,
        [string]$Name,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $typeValue = Resolve-EnumValue -Map $script:PPT_SHAPE_TYPE -Key $ShapeType

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex", "Add $ShapeType shape")) { return }

    $shape = $slide.Shapes.AddShape($typeValue, $Left, $Top, $Width, $Height)
    if ($Name) { $shape.Name = $Name }

    $result = [ordered]@{
        status = 'added'
        name   = $shape.Name
        index  = $shape.ZOrderPosition
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointTextBox
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointTextBox {
    <#
    .SYNOPSIS
        Add a text box to a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER Left
        Left position in points.
    .PARAMETER Top
        Top position in points.
    .PARAMETER Width
        Width in points.
    .PARAMETER Height
        Height in points.
    .PARAMETER Text
        Optional initial text content.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointTextBox -SlideIndex 1 -Left 50 -Top 50 -Width 300 -Height 40 -Text "Hello" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][double]$Left,
        [Parameter(Mandatory)][double]$Top,
        [Parameter(Mandatory)][double]$Width,
        [Parameter(Mandatory)][double]$Height,
        [string]$Text,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex", "Add text box")) { return }

    # 1 = msoTextOrientationHorizontal
    $shape = $slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    if ($Text) { $shape.TextFrame.TextRange.Text = $Text }

    $result = [ordered]@{
        status = 'added'
        name   = $shape.Name
        left   = $shape.Left
        top    = $shape.Top
        width  = $shape.Width
        height = $shape.Height
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointLine
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointLine {
    <#
    .SYNOPSIS
        Add a line shape to a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER BeginX
        Starting X coordinate in points.
    .PARAMETER BeginY
        Starting Y coordinate in points.
    .PARAMETER EndX
        Ending X coordinate in points.
    .PARAMETER EndY
        Ending Y coordinate in points.
    .PARAMETER Name
        Optional name for the line shape.
    .PARAMETER Weight
        Line weight in points.
    .PARAMETER Color
        Line color as 'R,G,B' string (e.g. '255,0,0').
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointLine -SlideIndex 1 -BeginX 10 -BeginY 10 -EndX 200 -EndY 200 -Weight 2 -Color "0,0,255" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][double]$BeginX,
        [Parameter(Mandatory)][double]$BeginY,
        [Parameter(Mandatory)][double]$EndX,
        [Parameter(Mandatory)][double]$EndY,
        [string]$Name,
        [double]$Weight,
        [string]$Color,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex", "Add line")) { return }

    $shape = $slide.Shapes.AddLine($BeginX, $BeginY, $EndX, $EndY)
    if ($Name)   { $shape.Name = $Name }
    if ($PSBoundParameters.ContainsKey('Weight')) { $shape.Line.Weight = $Weight }
    if ($Color)  { $shape.Line.ForeColor.RGB = ConvertFrom-RGBString $Color }

    $result = [ordered]@{
        status = 'added'
        name   = $shape.Name
        left   = $shape.Left
        top    = $shape.Top
        width  = $shape.Width
        height = $shape.Height
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Remove-PowerPointShape
# ══════════════════════════════════════════════════════════════════════════

function Remove-PowerPointShape {
    <#
    .SYNOPSIS
        Delete a shape from a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape to delete.
    .PARAMETER ShapeIndex
        1-based index of the shape to delete.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-PowerPointShape -SlideIndex 1 -ShapeName "Rectangle 1" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
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

    $deletedName = $shape.Name
    if (-not $PSCmdlet.ShouldProcess("Shape '$deletedName' on slide $SlideIndex", "Delete")) { return }

    $shape.Delete()

    $result = [ordered]@{
        status       = 'deleted'
        deletedShape = $deletedName
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointShapeProperties
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointShapeProperties {
    <#
    .SYNOPSIS
        Modify properties of a shape on a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape to modify.
    .PARAMETER ShapeIndex
        1-based index of the shape to modify.
    .PARAMETER Name
        New name for the shape.
    .PARAMETER Left
        New left position in points.
    .PARAMETER Top
        New top position in points.
    .PARAMETER Width
        New width in points.
    .PARAMETER Height
        New height in points.
    .PARAMETER Rotation
        Rotation angle in degrees.
    .PARAMETER Visible
        Shape visibility.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointShapeProperties -SlideIndex 1 -ShapeName "Title 1" -Left 50 -Top 50 -Width 400 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [string]$Name,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [double]$Rotation,
        [bool]$Visible,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = $ShapeName
    if (-not $target) { $target = "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Set properties")) { return }

    $changed = @()
    if ($PSBoundParameters.ContainsKey('Name'))     { $shape.Name     = $Name;     $changed += 'Name' }
    if ($PSBoundParameters.ContainsKey('Left'))     { $shape.Left     = $Left;     $changed += 'Left' }
    if ($PSBoundParameters.ContainsKey('Top'))      { $shape.Top      = $Top;      $changed += 'Top' }
    if ($PSBoundParameters.ContainsKey('Width'))    { $shape.Width    = $Width;    $changed += 'Width' }
    if ($PSBoundParameters.ContainsKey('Height'))   { $shape.Height   = $Height;   $changed += 'Height' }
    if ($PSBoundParameters.ContainsKey('Rotation')) { $shape.Rotation = $Rotation; $changed += 'Rotation' }
    if ($PSBoundParameters.ContainsKey('Visible'))  {
        $shape.Visible = if ($Visible) { -1 } else { 0 }  # msoTrue / msoFalse
        $changed += 'Visible'
    }

    $result = [ordered]@{
        status  = 'modified'
        shape   = $shape.Name
        changed = $changed
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Copy-PowerPointShape
# ══════════════════════════════════════════════════════════════════════════

function Copy-PowerPointShape {
    <#
    .SYNOPSIS
        Copy a shape, optionally to a different slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based source slide index.
    .PARAMETER ShapeName
        Name of the shape to copy.
    .PARAMETER ShapeIndex
        1-based index of the shape to copy.
    .PARAMETER TargetSlideIndex
        1-based target slide index. Defaults to the same slide.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Copy-PowerPointShape -SlideIndex 1 -ShapeName "Logo" -TargetSlideIndex 2 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [int]$TargetSlideIndex,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $targetIdx   = if ($PSBoundParameters.ContainsKey('TargetSlideIndex')) { $TargetSlideIndex } else { $SlideIndex }
    $targetSlide = $pres.Slides.Item($targetIdx)

    $shape.Copy()
    $pasted = $targetSlide.Shapes.Paste()
    $newShape = $pasted.Item(1)

    $result = [ordered]@{
        status     = 'copied'
        name       = $newShape.Name
        slideIndex = $targetIdx
        left       = $newShape.Left
        top        = $newShape.Top
        width      = $newShape.Width
        height     = $newShape.Height
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Group-PowerPointShapes
# ══════════════════════════════════════════════════════════════════════════

function Group-PowerPointShapes {
    <#
    .SYNOPSIS
        Group multiple shapes on a slide into a single group shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeNames
        Array of shape names to group.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Group-PowerPointShapes -SlideIndex 1 -ShapeNames @("Rect1","Rect2","Oval1") -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string[]]$ShapeNames,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $grouped = $slide.Shapes.Range($ShapeNames).Group()

    $result = [ordered]@{
        status = 'grouped'
        name   = $grouped.Name
        left   = $grouped.Left
        top    = $grouped.Top
        width  = $grouped.Width
        height = $grouped.Height
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Ungroup-PowerPointShapes
# ══════════════════════════════════════════════════════════════════════════

function Ungroup-PowerPointShapes {
    <#
    .SYNOPSIS
        Ungroup a grouped shape into its individual components.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the grouped shape to ungroup.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Ungroup-PowerPointShapes -SlideIndex 1 -ShapeName "Group 1" -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$ShapeName,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = $slide.Shapes.Item($ShapeName)

    $ungrouped = $shape.Ungroup()
    $names = @()
    foreach ($s in $ungrouped) { $names += $s.Name }

    $result = [ordered]@{
        status = 'ungrouped'
        shapes = $names
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointShapeZOrder
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointShapeZOrder {
    <#
    .SYNOPSIS
        Change the z-order of a shape on a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER ZOrderCmd
        Z-order command key (e.g. 'bringToFront', 'sendToBack'). Resolved via PPT_ZORDER map.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointShapeZOrder -SlideIndex 1 -ShapeName "Rect1" -ZOrderCmd bringToFront -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][string]$ZOrderCmd,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $cmdValue = Resolve-EnumValue -Map $script:PPT_ZORDER -Key $ZOrderCmd
    $target   = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "ZOrder $ZOrderCmd")) { return }

    $shape.ZOrder($cmdValue)

    $result = [ordered]@{
        status         = 'modified'
        shape          = $shape.Name
        zOrderCmd      = $ZOrderCmd
        zOrderPosition = $shape.ZOrderPosition
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointShapeList
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointShapeList {
    <#
    .SYNOPSIS
        List all shapes on a slide with basic info.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointShapeList -SlideIndex 1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $result = @()
    foreach ($s in $slide.Shapes) {
        $result += [ordered]@{
            name   = $s.Name
            type   = [int]$s.Type
            left   = $s.Left
            top    = $s.Top
            width  = $s.Width
            height = $s.Height
        }
    }

    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Find-PowerPointShape
# ══════════════════════════════════════════════════════════════════════════

function Find-PowerPointShape {
    <#
    .SYNOPSIS
        Search all slides for shapes by name or type.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER ShapeName
        Shape name to search for (case-insensitive substring match).
    .PARAMETER ShapeType
        Shape type key to filter by (e.g. 'rectangle'). Resolved via PPT_SHAPE_TYPE map.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Find-PowerPointShape -ShapeName "Logo" -AsJson
    .EXAMPLE
        Find-PowerPointShape -ShapeType oval -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [string]$ShapeName,
        [string]$ShapeType,
        [switch]$AsJson
    )

    if (-not $ShapeName -and -not $ShapeType) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentException]::new("Either -ShapeName or -ShapeType must be specified."),
            'MissingSearchCriteria',
            [System.Management.Automation.ErrorCategory]::InvalidArgument, $null)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app  = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres = $app.ActivePresentation

    $typeValue = if ($ShapeType) { Resolve-EnumValue -Map $script:PPT_SHAPE_TYPE -Key $ShapeType } else { $null }

    $result = @()
    foreach ($slide in $pres.Slides) {
        foreach ($s in $slide.Shapes) {
            $match = $true
            if ($ShapeName -and $s.Name -notlike "*$ShapeName*") { $match = $false }
            if ($null -ne $typeValue -and [int]$s.Type -ne $typeValue) { $match = $false }
            if ($match) {
                $result += [ordered]@{
                    slideIndex = $slide.SlideIndex
                    shapeName  = $s.Name
                    shapeType  = [int]$s.Type
                    left       = $s.Left
                    top        = $s.Top
                }
            }
        }
    }

    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Align-PowerPointShapes
# ══════════════════════════════════════════════════════════════════════════

function Align-PowerPointShapes {
    <#
    .SYNOPSIS
        Align multiple shapes on a slide relative to each other.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeNames
        Array of shape names to align.
    .PARAMETER AlignCmd
        Alignment command key (e.g. 'alignLefts', 'alignCenters'). Resolved via PPT_ALIGN_CMD map.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Align-PowerPointShapes -SlideIndex 1 -ShapeNames @("Rect1","Rect2") -AlignCmd alignCenters -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string[]]$ShapeNames,
        [Parameter(Mandatory)][string]$AlignCmd,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $cmdValue = Resolve-EnumValue -Map $script:PPT_ALIGN_CMD -Key $AlignCmd
    if (-not $PSCmdlet.ShouldProcess("Shapes on slide $SlideIndex", "Align $AlignCmd")) { return }

    # 0 = msoFalse (relative to each other, not slide)
    $slide.Shapes.Range($ShapeNames).Align($cmdValue, 0)

    $result = [ordered]@{
        status   = 'aligned'
        alignCmd = $AlignCmd
        shapes   = $ShapeNames
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Distribute-PowerPointShapes
# ══════════════════════════════════════════════════════════════════════════

function Distribute-PowerPointShapes {
    <#
    .SYNOPSIS
        Distribute multiple shapes evenly on a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeNames
        Array of shape names to distribute.
    .PARAMETER DistributeCmd
        Distribution command key (e.g. 'distributeHorizontally', 'distributeVertically'). Resolved via PPT_DISTRIBUTE map.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Distribute-PowerPointShapes -SlideIndex 1 -ShapeNames @("Rect1","Rect2","Rect3") -DistributeCmd distributeHorizontally -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string[]]$ShapeNames,
        [Parameter(Mandatory)][string]$DistributeCmd,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $cmdValue = Resolve-EnumValue -Map $script:PPT_DISTRIBUTE -Key $DistributeCmd
    if (-not $PSCmdlet.ShouldProcess("Shapes on slide $SlideIndex", "Distribute $DistributeCmd")) { return }

    # 0 = msoFalse (relative to each other, not slide)
    $slide.Shapes.Range($ShapeNames).Distribute($cmdValue, 0)

    $result = [ordered]@{
        status        = 'distributed'
        distributeCmd = $DistributeCmd
        shapes        = $ShapeNames
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
