# Public/FormattingOps.ps1 — Shape fill, line, shadow, effect, and positioning operations

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointShapeFill
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointShapeFill {
    <#
    .SYNOPSIS
        Set the fill style of a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER FillType
        Fill type: 'solid', 'gradient', 'picture', or 'none'.
    .PARAMETER Color
        Fill color as 'R,G,B' string. Used with 'solid' fill type.
    .PARAMETER Color1
        First gradient color as 'R,G,B' string. Used with 'gradient' fill type.
    .PARAMETER Color2
        Second gradient color as 'R,G,B' string. Used with 'gradient' fill type.
    .PARAMETER ImagePath
        Path to image file. Used with 'picture' fill type.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointShapeFill -SlideIndex 1 -ShapeName "Rect1" -FillType solid -Color "0,120,215" -AsJson
    .EXAMPLE
        Set-PowerPointShapeFill -SlideIndex 1 -ShapeName "Rect1" -FillType gradient -Color1 "255,0,0" -Color2 "0,0,255" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][ValidateSet('solid','gradient','picture','none')][string]$FillType,
        [string]$Color,
        [string]$Color1,
        [string]$Color2,
        [string]$ImagePath,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Set fill to $FillType")) { return }

    $fill = $shape.Fill

    switch ($FillType.ToLower()) {
        'solid' {
            if (-not $Color) {
                $er = [System.Management.Automation.ErrorRecord]::new(
                    [System.ArgumentException]::new("-Color is required for solid fill."),
                    'MissingColor', [System.Management.Automation.ErrorCategory]::InvalidArgument, $null)
                $PSCmdlet.ThrowTerminatingError($er)
            }
            $fill.Solid()
            $fill.ForeColor.RGB = ConvertFrom-RGBString $Color
        }
        'gradient' {
            if (-not $Color1 -or -not $Color2) {
                $er = [System.Management.Automation.ErrorRecord]::new(
                    [System.ArgumentException]::new("-Color1 and -Color2 are required for gradient fill."),
                    'MissingGradientColors', [System.Management.Automation.ErrorCategory]::InvalidArgument, $null)
                $PSCmdlet.ThrowTerminatingError($er)
            }
            # 1 = msoGradientHorizontal, variant 1
            $fill.TwoColorGradient(1, 1)
            $fill.ForeColor.RGB = ConvertFrom-RGBString $Color1
            $fill.BackColor.RGB = ConvertFrom-RGBString $Color2
        }
        'picture' {
            if (-not $ImagePath) {
                $er = [System.Management.Automation.ErrorRecord]::new(
                    [System.ArgumentException]::new("-ImagePath is required for picture fill."),
                    'MissingImagePath', [System.Management.Automation.ErrorCategory]::InvalidArgument, $null)
                $PSCmdlet.ThrowTerminatingError($er)
            }
            if (-not (Test-Path -LiteralPath $ImagePath)) {
                $er = [System.Management.Automation.ErrorRecord]::new(
                    [System.IO.FileNotFoundException]::new("Image file not found: $ImagePath"),
                    'ImageNotFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $ImagePath)
                $PSCmdlet.ThrowTerminatingError($er)
            }
            $fill.UserPicture([System.IO.Path]::GetFullPath($ImagePath))
        }
        'none' {
            $fill.Visible = 0  # msoFalse
        }
    }

    $result = [ordered]@{
        status   = 'modified'
        shape    = $shape.Name
        fillType = $FillType
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointShapeLine
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointShapeLine {
    <#
    .SYNOPSIS
        Set line (border) properties of a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER Color
        Line color as 'R,G,B' string.
    .PARAMETER Weight
        Line weight in points.
    .PARAMETER DashStyle
        Dash style key (e.g. 'solid', 'dash', 'roundDot'). Resolved via PPT_LINE_DASH map.
    .PARAMETER Visible
        Line visibility.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointShapeLine -SlideIndex 1 -ShapeName "Rect1" -Color "0,0,0" -Weight 2 -DashStyle dash -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [string]$Color,
        [double]$Weight,
        [string]$DashStyle,
        [bool]$Visible,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Set line properties")) { return }

    $line    = $shape.Line
    $changed = @()

    if ($PSBoundParameters.ContainsKey('Color'))     { $line.ForeColor.RGB = ConvertFrom-RGBString $Color; $changed += 'Color' }
    if ($PSBoundParameters.ContainsKey('Weight'))    { $line.Weight    = $Weight;                          $changed += 'Weight' }
    if ($PSBoundParameters.ContainsKey('DashStyle')) { $line.DashStyle = Resolve-EnumValue -Map $script:PPT_LINE_DASH -Key $DashStyle; $changed += 'DashStyle' }
    if ($PSBoundParameters.ContainsKey('Visible'))   { $line.Visible   = if ($Visible) { -1 } else { 0 }; $changed += 'Visible' }

    $result = [ordered]@{
        status  = 'modified'
        shape   = $shape.Name
        changed = $changed
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointShapeShadow
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointShapeShadow {
    <#
    .SYNOPSIS
        Configure shadow effect on a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER Visible
        Shadow visibility.
    .PARAMETER OffsetX
        Horizontal shadow offset in points.
    .PARAMETER OffsetY
        Vertical shadow offset in points.
    .PARAMETER Blur
        Shadow blur radius in points.
    .PARAMETER Color
        Shadow color as 'R,G,B' string.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointShapeShadow -SlideIndex 1 -ShapeName "Rect1" -Visible $true -OffsetX 3 -OffsetY 3 -Blur 5 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][bool]$Visible,
        [double]$OffsetX,
        [double]$OffsetY,
        [double]$Blur,
        [string]$Color,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Set shadow")) { return }

    $shadow  = $shape.Shadow
    $changed = @()

    $shadow.Visible = if ($Visible) { -1 } else { 0 }
    $changed += 'Visible'

    if ($PSBoundParameters.ContainsKey('OffsetX')) { $shadow.OffsetX = $OffsetX; $changed += 'OffsetX' }
    if ($PSBoundParameters.ContainsKey('OffsetY')) { $shadow.OffsetY = $OffsetY; $changed += 'OffsetY' }
    if ($PSBoundParameters.ContainsKey('Blur'))    { $shadow.Blur    = $Blur;    $changed += 'Blur' }
    if ($PSBoundParameters.ContainsKey('Color'))   { $shadow.ForeColor.RGB = ConvertFrom-RGBString $Color; $changed += 'Color' }

    $result = [ordered]@{
        status  = 'modified'
        shape   = $shape.Name
        changed = $changed
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointShapeEffect
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointShapeEffect {
    <#
    .SYNOPSIS
        Apply glow, soft edge, or reflection effects to a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER GlowRadius
        Glow effect radius in points.
    .PARAMETER GlowColor
        Glow color as 'R,G,B' string.
    .PARAMETER SoftEdgeRadius
        Soft edge radius in points.
    .PARAMETER ReflectionType
        Reflection type as integer (0=none, 1-9 for various reflection styles).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointShapeEffect -SlideIndex 1 -ShapeName "Rect1" -GlowRadius 10 -GlowColor "255,215,0" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [double]$GlowRadius,
        [string]$GlowColor,
        [double]$SoftEdgeRadius,
        [int]$ReflectionType,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Set effects")) { return }

    $changed = @()

    if ($PSBoundParameters.ContainsKey('GlowRadius')) {
        $shape.Glow.Radius = $GlowRadius
        $changed += 'GlowRadius'
    }
    if ($PSBoundParameters.ContainsKey('GlowColor')) {
        $shape.Glow.Color.RGB = ConvertFrom-RGBString $GlowColor
        $changed += 'GlowColor'
    }
    if ($PSBoundParameters.ContainsKey('SoftEdgeRadius')) {
        $shape.SoftEdge.Radius = $SoftEdgeRadius
        $changed += 'SoftEdgeRadius'
    }
    if ($PSBoundParameters.ContainsKey('ReflectionType')) {
        $shape.Reflection.Type = $ReflectionType
        $changed += 'ReflectionType'
    }

    $result = [ordered]@{
        status  = 'modified'
        shape   = $shape.Name
        changed = $changed
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointThemeColor
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointThemeColor {
    <#
    .SYNOPSIS
        Apply an Office theme color to a shape's fill.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER ThemeColor
        Theme color index (1-12). Maps to MsoThemeColorIndex.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointThemeColor -SlideIndex 1 -ShapeName "Rect1" -ThemeColor 5 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][ValidateRange(1,12)][int]$ThemeColor,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Set theme color $ThemeColor")) { return }

    $shape.Fill.Solid()
    $shape.Fill.ForeColor.ObjectThemeColor = $ThemeColor

    $result = [ordered]@{
        status     = 'modified'
        shape      = $shape.Name
        themeColor = $ThemeColor
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointShapeSize
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointShapeSize {
    <#
    .SYNOPSIS
        Set the width and height of a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER Width
        New width in points.
    .PARAMETER Height
        New height in points.
    .PARAMETER LockAspectRatio
        Lock the aspect ratio when resizing.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointShapeSize -SlideIndex 1 -ShapeName "Rect1" -Width 300 -Height 200 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][double]$Width,
        [Parameter(Mandatory)][double]$Height,
        [bool]$LockAspectRatio,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Set size ${Width}x${Height}")) { return }

    if ($PSBoundParameters.ContainsKey('LockAspectRatio')) {
        $shape.LockAspectRatio = if ($LockAspectRatio) { -1 } else { 0 }  # msoTrue / msoFalse
    }

    $shape.Width  = $Width
    $shape.Height = $Height

    $result = [ordered]@{
        status = 'modified'
        shape  = $shape.Name
        width  = $shape.Width
        height = $shape.Height
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointShapePosition
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointShapePosition {
    <#
    .SYNOPSIS
        Set the position of a shape on a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER Left
        Left position in points.
    .PARAMETER Top
        Top position in points.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointShapePosition -SlideIndex 1 -ShapeName "Rect1" -Left 100 -Top 50 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][double]$Left,
        [Parameter(Mandatory)][double]$Top,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Set position ($Left, $Top)")) { return }

    $shape.Left = $Left
    $shape.Top  = $Top

    $result = [ordered]@{
        status = 'modified'
        shape  = $shape.Name
        left   = $shape.Left
        top    = $shape.Top
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
