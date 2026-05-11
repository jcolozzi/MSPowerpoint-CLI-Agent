# Public/AnimationOps.ps1 — Animation operations

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointAnimation
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointAnimation {
    <#
    .SYNOPSIS
        Add an animation effect to a shape on a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape to animate.
    .PARAMETER ShapeIndex
        1-based index of the shape to animate.
    .PARAMETER EffectId
        Animation effect key (e.g. 'fadeIn', 'flyIn'). Resolved via PPT_ANIM_EFFECT map.
    .PARAMETER Duration
        Animation duration in seconds. Default 1.
    .PARAMETER Delay
        Delay before animation starts in seconds. Default 0.
    .PARAMETER TriggerType
        When the animation triggers: 'onClick' (default), 'withPrevious', 'afterPrevious'.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointAnimation -SlideIndex 1 -ShapeName "Title 1" -EffectId fadeIn -Duration 0.5 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][string]$EffectId,
        [double]$Duration = 1,
        [double]$Delay = 0,
        [ValidateSet('onClick','withPrevious','afterPrevious')]
        [string]$TriggerType = 'onClick',
        [switch]$AsJson
    )

    $triggerMap = @{ onClick = 1; withPrevious = 2; afterPrevious = 3 }

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $shape       = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex
    $effectValue = Resolve-EnumValue -Map $script:PPT_ANIM_EFFECT -Key $EffectId
    $triggerValue = $triggerMap[$TriggerType]

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex, Shape '$($shape.Name)'", "Add animation '$EffectId'")) { return }

    $effect = $slide.TimeLine.MainSequence.AddEffect($shape, $effectValue)
    $effect.Timing.Duration         = $Duration
    $effect.Timing.TriggerDelayTime = $Delay
    $effect.Timing.TriggerType      = $triggerValue

    $result = [ordered]@{
        status      = 'added'
        effectIndex = $effect.Index
        effectId    = $EffectId
        shapeName   = $shape.Name
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointAnimation
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointAnimation {
    <#
    .SYNOPSIS
        Get animation effects from a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER EffectIndex
        1-based effect index. If specified, returns a single effect. Otherwise returns all.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointAnimation -SlideIndex 1 -AsJson
    .EXAMPLE
        Get-PowerPointAnimation -SlideIndex 2 -EffectIndex 1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [int]$EffectIndex,
        [switch]$AsJson
    )

    $triggerNames = @{ 1 = 'onClick'; 2 = 'withPrevious'; 3 = 'afterPrevious' }

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $getEffectInfo = {
        param($e, $idx)
        [ordered]@{
            index       = $idx
            shapeName   = $e.Shape.Name
            effectType  = [int]$e.EffectType
            duration    = $e.Timing.Duration
            delay       = $e.Timing.TriggerDelayTime
            triggerType = if ($triggerNames.ContainsKey([int]$e.Timing.TriggerType)) { $triggerNames[[int]$e.Timing.TriggerType] } else { [int]$e.Timing.TriggerType }
        }
    }

    $seq = $slide.TimeLine.MainSequence

    if ($EffectIndex) {
        $effect = $seq.Item($EffectIndex)
        $result = & $getEffectInfo $effect $EffectIndex
    }
    else {
        $result = @()
        for ($i = 1; $i -le $seq.Count; $i++) {
            $result += & $getEffectInfo $seq.Item($i) $i
        }
    }

    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Remove-PowerPointAnimation
# ══════════════════════════════════════════════════════════════════════════

function Remove-PowerPointAnimation {
    <#
    .SYNOPSIS
        Remove a specific animation effect from a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER EffectIndex
        1-based index of the effect to remove.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-PowerPointAnimation -SlideIndex 1 -EffectIndex 2 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][int]$EffectIndex,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex, Effect $EffectIndex", 'Remove animation')) { return }

    $slide.TimeLine.MainSequence.Item($EffectIndex).Delete()

    $result = [ordered]@{ status = 'removed'; effectIndex = $EffectIndex }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointAnimationTiming
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointAnimationTiming {
    <#
    .SYNOPSIS
        Modify timing properties of an existing animation effect.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER EffectIndex
        1-based effect index.
    .PARAMETER Duration
        Animation duration in seconds.
    .PARAMETER Delay
        Delay before animation starts in seconds.
    .PARAMETER TriggerType
        When the animation triggers: 'onClick', 'withPrevious', 'afterPrevious'.
    .PARAMETER RepeatCount
        Number of times to repeat the animation.
    .PARAMETER Rewind
        Whether the animation rewinds after playing.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointAnimationTiming -SlideIndex 1 -EffectIndex 1 -Duration 2.5 -TriggerType afterPrevious -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][int]$EffectIndex,
        [double]$Duration,
        [double]$Delay,
        [ValidateSet('onClick','withPrevious','afterPrevious')]
        [string]$TriggerType,
        [int]$RepeatCount,
        [bool]$Rewind,
        [switch]$AsJson
    )

    $triggerMap = @{ onClick = 1; withPrevious = 2; afterPrevious = 3 }

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $effect = $slide.TimeLine.MainSequence.Item($EffectIndex)

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex, Effect $EffectIndex", 'Set animation timing')) { return }

    $timing = $effect.Timing
    if ($PSBoundParameters.ContainsKey('Duration'))    { $timing.Duration         = $Duration }
    if ($PSBoundParameters.ContainsKey('Delay'))        { $timing.TriggerDelayTime = $Delay }
    if ($PSBoundParameters.ContainsKey('TriggerType'))  { $timing.TriggerType      = $triggerMap[$TriggerType] }
    if ($PSBoundParameters.ContainsKey('RepeatCount'))  { $timing.RepeatCount      = $RepeatCount }
    if ($PSBoundParameters.ContainsKey('Rewind'))       { $timing.Rewind           = if ($Rewind) { -1 } else { 0 } }

    $result = [ordered]@{ status = 'updated'; effectIndex = $EffectIndex }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointAnimationOrder
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointAnimationOrder {
    <#
    .SYNOPSIS
        Change the playback order of an animation effect in the sequence.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER EffectIndex
        1-based current index of the effect.
    .PARAMETER NewIndex
        1-based desired new index for the effect.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointAnimationOrder -SlideIndex 1 -EffectIndex 3 -NewIndex 1 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][int]$EffectIndex,
        [Parameter(Mandatory)][int]$NewIndex,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex, Effect $EffectIndex → $NewIndex", 'Reorder animation')) { return }

    $effect = $slide.TimeLine.MainSequence.Item($EffectIndex)
    if ($NewIndex -lt $EffectIndex) {
        $effect.MoveBefore($NewIndex)
    }
    else {
        $effect.MoveAfter($NewIndex)
    }

    $result = [ordered]@{ status = 'reordered'; fromIndex = $EffectIndex; toIndex = $NewIndex }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Clear-PowerPointAnimations
# ══════════════════════════════════════════════════════════════════════════

function Clear-PowerPointAnimations {
    <#
    .SYNOPSIS
        Remove all animation effects from a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Clear-PowerPointAnimations -SlideIndex 1 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $seq   = $slide.TimeLine.MainSequence
    $count = $seq.Count

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex ($count effects)", 'Clear all animations')) { return }

    for ($i = $count; $i -ge 1; $i--) {
        $seq.Item($i).Delete()
    }

    $result = [ordered]@{ status = 'cleared'; removedCount = $count }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Copy-PowerPointAnimation
# ══════════════════════════════════════════════════════════════════════════

function Copy-PowerPointAnimation {
    <#
    .SYNOPSIS
        Copy an animation effect from one slide/shape to another.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based source slide index.
    .PARAMETER EffectIndex
        1-based source effect index.
    .PARAMETER TargetSlideIndex
        1-based target slide index.
    .PARAMETER TargetShapeName
        Name of the target shape.
    .PARAMETER TargetShapeIndex
        1-based index of the target shape.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Copy-PowerPointAnimation -SlideIndex 1 -EffectIndex 1 -TargetSlideIndex 2 -TargetShapeName "Title 1" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][int]$EffectIndex,
        [Parameter(Mandatory)][int]$TargetSlideIndex,
        [string]$TargetShapeName,
        [int]$TargetShapeIndex,
        [switch]$AsJson
    )

    $triggerMap = @{ 1 = 'onClick'; 2 = 'withPrevious'; 3 = 'afterPrevious' }

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation

    # Read source effect properties
    $srcSlide  = $pres.Slides.Item($SlideIndex)
    $srcEffect = $srcSlide.TimeLine.MainSequence.Item($EffectIndex)

    $effectType  = [int]$srcEffect.EffectType
    $duration    = $srcEffect.Timing.Duration
    $delay       = $srcEffect.Timing.TriggerDelayTime
    $triggerType = [int]$srcEffect.Timing.TriggerType

    # Get target shape
    $tgtSlide = $pres.Slides.Item($TargetSlideIndex)
    $tgtShape = Get-SlideShape -Slide $tgtSlide -ShapeName $TargetShapeName -ShapeIndex $TargetShapeIndex

    if (-not $PSCmdlet.ShouldProcess("Slide $TargetSlideIndex, Shape '$($tgtShape.Name)'", "Copy animation from Slide $SlideIndex Effect $EffectIndex")) { return }

    $newEffect = $tgtSlide.TimeLine.MainSequence.AddEffect($tgtShape, $effectType)
    $newEffect.Timing.Duration         = $duration
    $newEffect.Timing.TriggerDelayTime = $delay
    $newEffect.Timing.TriggerType      = $triggerType

    $result = [ordered]@{
        status           = 'copied'
        sourceSlide      = $SlideIndex
        sourceEffect     = $EffectIndex
        targetSlide      = $TargetSlideIndex
        targetShapeName  = $tgtShape.Name
        newEffectIndex   = $newEffect.Index
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
