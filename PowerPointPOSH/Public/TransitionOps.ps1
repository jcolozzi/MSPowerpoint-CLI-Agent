# Public/TransitionOps.ps1 — Slide transition operations

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointTransition
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointTransition {
    <#
    .SYNOPSIS
        Set a slide transition effect on one or all slides.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index, or 'all' to apply to every slide.
    .PARAMETER TransitionType
        Transition type key (e.g. 'fade', 'dissolve'). Resolved via PPT_TRANSITION map.
    .PARAMETER Duration
        Transition duration in seconds. Default 1.
    .PARAMETER AdvanceOnClick
        Whether clicking advances the slide. Default $true.
    .PARAMETER AdvanceAfterTime
        Auto-advance time in seconds. 0 disables auto-advance.
    .PARAMETER Sound
        Path to a .wav file to play during the transition.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointTransition -SlideIndex 1 -TransitionType fade -Duration 1.5 -AsJson
    .EXAMPLE
        Set-PowerPointTransition -SlideIndex all -TransitionType dissolve -AdvanceAfterTime 5 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$SlideIndex,
        [Parameter(Mandatory)][string]$TransitionType,
        [double]$Duration = 1,
        [bool]$AdvanceOnClick = $true,
        [double]$AdvanceAfterTime = 0,
        [string]$Sound,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation

    $transValue = Resolve-EnumValue -Map $script:PPT_TRANSITION -Key $TransitionType

    # Build list of slides to process
    if ($SlideIndex -ieq 'all') {
        $slides = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) { $slides += $pres.Slides.Item($i) }
        $target = "All slides"
    }
    else {
        $slides = @($pres.Slides.Item([int]$SlideIndex))
        $target = "Slide $SlideIndex"
    }

    if (-not $PSCmdlet.ShouldProcess($target, "Set transition '$TransitionType'")) { return }

    $applied = 0
    foreach ($sl in $slides) {
        $trans = $sl.SlideShowTransition
        $trans.EntryEffect    = $transValue
        $trans.Duration       = $Duration
        $trans.AdvanceOnClick = if ($AdvanceOnClick) { -1 } else { 0 }

        if ($AdvanceAfterTime -gt 0) {
            $trans.AdvanceOnTime = -1          # msoTrue
            $trans.AdvanceTime   = $AdvanceAfterTime
        }
        else {
            $trans.AdvanceOnTime = 0           # msoFalse
        }

        if ($Sound) {
            $trans.SoundEffect.ImportFromFile($Sound)
        }
        $applied++
    }

    $result = [ordered]@{ status = 'set'; transitionType = $TransitionType; slidesAffected = $applied }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointTransition
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointTransition {
    <#
    .SYNOPSIS
        Get transition settings for one or all slides.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index. If omitted, returns transitions for all slides.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointTransition -AsJson
    .EXAMPLE
        Get-PowerPointTransition -SlideIndex 3 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [int]$SlideIndex,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation

    # Reverse map: value → friendly name
    $reverseMap = @{}
    foreach ($entry in $script:PPT_TRANSITION.GetEnumerator()) {
        $reverseMap[[int]$entry.Value] = $entry.Key
    }

    $getTransInfo = {
        param($sl)
        $trans = $sl.SlideShowTransition
        $effectVal  = [int]$trans.EntryEffect
        $friendlyName = if ($reverseMap.ContainsKey($effectVal)) { $reverseMap[$effectVal] } else { "unknown($effectVal)" }
        [ordered]@{
            slideIndex     = $sl.SlideIndex
            entryEffect    = $friendlyName
            entryEffectId  = $effectVal
            duration       = $trans.Duration
            advanceOnClick = $trans.AdvanceOnClick -eq -1
            advanceOnTime  = $trans.AdvanceOnTime -eq -1
            advanceTime    = $trans.AdvanceTime
        }
    }

    if ($SlideIndex) {
        $result = & $getTransInfo $pres.Slides.Item($SlideIndex)
    }
    else {
        $result = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $result += & $getTransInfo $pres.Slides.Item($i)
        }
    }

    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Remove-PowerPointTransition
# ══════════════════════════════════════════════════════════════════════════

function Remove-PowerPointTransition {
    <#
    .SYNOPSIS
        Remove the transition effect from one or all slides.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index, or 'all' to clear transitions from every slide.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-PowerPointTransition -SlideIndex 1 -AsJson
    .EXAMPLE
        Remove-PowerPointTransition -SlideIndex all -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$SlideIndex,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation

    if ($SlideIndex -ieq 'all') {
        $slides = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) { $slides += $pres.Slides.Item($i) }
        $target = "All slides"
    }
    else {
        $slides = @($pres.Slides.Item([int]$SlideIndex))
        $target = "Slide $SlideIndex"
    }

    if (-not $PSCmdlet.ShouldProcess($target, 'Remove transition')) { return }

    $removed = 0
    foreach ($sl in $slides) {
        $sl.SlideShowTransition.EntryEffect = 0   # ppEffectNone
        $removed++
    }

    $result = [ordered]@{ status = 'removed'; slidesAffected = $removed }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Copy-PowerPointTransition
# ══════════════════════════════════════════════════════════════════════════

function Copy-PowerPointTransition {
    <#
    .SYNOPSIS
        Copy transition settings from one slide to another (or all).
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SourceSlideIndex
        1-based source slide index.
    .PARAMETER TargetSlideIndex
        1-based target slide index, or 'all' to apply to every slide.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Copy-PowerPointTransition -SourceSlideIndex 1 -TargetSlideIndex 3 -AsJson
    .EXAMPLE
        Copy-PowerPointTransition -SourceSlideIndex 1 -TargetSlideIndex all -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SourceSlideIndex,
        [Parameter(Mandatory)][string]$TargetSlideIndex,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation

    # Read source transition
    $srcTrans = $pres.Slides.Item($SourceSlideIndex).SlideShowTransition
    $srcEffect       = [int]$srcTrans.EntryEffect
    $srcDuration     = $srcTrans.Duration
    $srcAdvClick     = [int]$srcTrans.AdvanceOnClick
    $srcAdvOnTime    = [int]$srcTrans.AdvanceOnTime
    $srcAdvTime      = $srcTrans.AdvanceTime

    # Build target list
    if ($TargetSlideIndex -ieq 'all') {
        $targets = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            if ($i -ne $SourceSlideIndex) { $targets += $pres.Slides.Item($i) }
        }
        $targetLabel = "All slides (except source $SourceSlideIndex)"
    }
    else {
        $targets = @($pres.Slides.Item([int]$TargetSlideIndex))
        $targetLabel = "Slide $TargetSlideIndex"
    }

    if (-not $PSCmdlet.ShouldProcess($targetLabel, "Copy transition from Slide $SourceSlideIndex")) { return }

    $applied = 0
    foreach ($sl in $targets) {
        $t = $sl.SlideShowTransition
        $t.EntryEffect    = $srcEffect
        $t.Duration       = $srcDuration
        $t.AdvanceOnClick = $srcAdvClick
        $t.AdvanceOnTime  = $srcAdvOnTime
        $t.AdvanceTime    = $srcAdvTime
        $applied++
    }

    $result = [ordered]@{
        status         = 'copied'
        sourceSlide    = $SourceSlideIndex
        slidesAffected = $applied
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
