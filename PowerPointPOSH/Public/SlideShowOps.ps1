# Public/SlideShowOps.ps1 — Slide show operations

# ══════════════════════════════════════════════════════════════════════════
# Start-PowerPointSlideShow
# ══════════════════════════════════════════════════════════════════════════

function Start-PowerPointSlideShow {
    <#
    .SYNOPSIS
        Start a slide show for the active presentation.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER StartSlide
        1-based starting slide index. Default 1.
    .PARAMETER EndSlide
        1-based ending slide index. Default is the last slide.
    .PARAMETER LoopUntilStopped
        Whether the show loops continuously.
    .PARAMETER WithNarration
        Whether narration plays. Default $true.
    .PARAMETER WithAnimation
        Whether animations play. Default $true.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Start-PowerPointSlideShow -AsJson
    .EXAMPLE
        Start-PowerPointSlideShow -StartSlide 3 -EndSlide 10 -LoopUntilStopped $true -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [int]$StartSlide = 1,
        [int]$EndSlide,
        [bool]$LoopUntilStopped = $false,
        [bool]$WithNarration = $true,
        [bool]$WithAnimation = $true,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation

    if (-not $EndSlide) { $EndSlide = $pres.Slides.Count }

    $settings = $pres.SlideShowSettings
    $settings.StartingSlide     = $StartSlide
    $settings.EndingSlide       = $EndSlide
    $settings.LoopUntilStopped  = if ($LoopUntilStopped) { -1 } else { 0 }
    $settings.ShowWithNarration = if ($WithNarration)     { -1 } else { 0 }
    $settings.ShowWithAnimation = if ($WithAnimation)     { -1 } else { 0 }

    $ssw = $settings.Run()

    $result = [ordered]@{
        status               = 'started'
        slideShowWindowIndex = $ssw.Presentation.Name
        startSlide           = $StartSlide
        endSlide             = $EndSlide
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Stop-PowerPointSlideShow
# ══════════════════════════════════════════════════════════════════════════

function Stop-PowerPointSlideShow {
    <#
    .SYNOPSIS
        Stop the currently running slide show.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Stop-PowerPointSlideShow -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app = Connect-PowerPointPresentation -PresentationPath $resolvedPath

    if ($app.SlideShowWindows.Count -gt 0) {
        $app.SlideShowWindows.Item(1).View.Exit()
        $status = 'stopped'
    }
    else {
        $status = 'notRunning'
    }

    $result = [ordered]@{ status = $status }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointSlideShowSettings
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointSlideShowSettings {
    <#
    .SYNOPSIS
        Configure slide show settings without starting the show.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER ShowType
        Presentation mode: 'speaker' (1), 'browsedByIndividual' (2), 'browsedAtKiosk' (3).
    .PARAMETER LoopUntilStopped
        Whether the show loops continuously.
    .PARAMETER ShowWithNarration
        Whether narration plays.
    .PARAMETER ShowWithAnimation
        Whether animations play.
    .PARAMETER AdvanceMode
        How slides advance: 'manually' (1) or 'useTimings' (2).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointSlideShowSettings -ShowType browsedAtKiosk -LoopUntilStopped $true -AsJson
    .EXAMPLE
        Set-PowerPointSlideShowSettings -AdvanceMode useTimings -ShowWithAnimation $false -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [ValidateSet('speaker','browsedByIndividual','browsedAtKiosk')]
        [string]$ShowType,
        [bool]$LoopUntilStopped,
        [bool]$ShowWithNarration,
        [bool]$ShowWithAnimation,
        [ValidateSet('manually','useTimings')]
        [string]$AdvanceMode,
        [switch]$AsJson
    )

    $showTypeMap   = @{ speaker = 1; browsedByIndividual = 2; browsedAtKiosk = 3 }
    $advanceModeMap = @{ manually = 1; useTimings = 2 }

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation

    if (-not $PSCmdlet.ShouldProcess('SlideShowSettings', 'Update settings')) { return }

    $settings = $pres.SlideShowSettings
    if ($PSBoundParameters.ContainsKey('ShowType'))          { $settings.ShowType          = $showTypeMap[$ShowType] }
    if ($PSBoundParameters.ContainsKey('LoopUntilStopped'))  { $settings.LoopUntilStopped  = if ($LoopUntilStopped) { -1 } else { 0 } }
    if ($PSBoundParameters.ContainsKey('ShowWithNarration')) { $settings.ShowWithNarration = if ($ShowWithNarration) { -1 } else { 0 } }
    if ($PSBoundParameters.ContainsKey('ShowWithAnimation')) { $settings.ShowWithAnimation = if ($ShowWithAnimation) { -1 } else { 0 } }
    if ($PSBoundParameters.ContainsKey('AdvanceMode'))       { $settings.AdvanceMode       = $advanceModeMap[$AdvanceMode] }

    $result = [ordered]@{ status = 'updated' }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointSlideShowInfo
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointSlideShowInfo {
    <#
    .SYNOPSIS
        Get information about the currently running slide show.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointSlideShowInfo -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [switch]$AsJson
    )

    $stateNames = @{ 1 = 'running'; 2 = 'paused'; 3 = 'blackScreen'; 4 = 'whiteScreen'; 5 = 'done' }

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app = Connect-PowerPointPresentation -PresentationPath $resolvedPath

    if ($app.SlideShowWindows.Count -gt 0) {
        $view     = $app.SlideShowWindows.Item(1).View
        $stateVal = [int]$view.State
        $result = [ordered]@{
            status       = 'running'
            currentSlide = $view.CurrentShowPosition
            totalSlides  = $app.ActivePresentation.Slides.Count
            state        = if ($stateNames.ContainsKey($stateVal)) { $stateNames[$stateVal] } else { "unknown($stateVal)" }
        }
    }
    else {
        $result = [ordered]@{ status = 'notRunning' }
    }

    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Step-PowerPointSlideShow
# ══════════════════════════════════════════════════════════════════════════

function Step-PowerPointSlideShow {
    <#
    .SYNOPSIS
        Navigate within a running slide show.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER Direction
        Navigation direction: 'next', 'previous', 'first', 'last', or a slide number.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Step-PowerPointSlideShow -Direction next -AsJson
    .EXAMPLE
        Step-PowerPointSlideShow -Direction 5 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$Direction,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app = Connect-PowerPointPresentation -PresentationPath $resolvedPath

    if ($app.SlideShowWindows.Count -eq 0) {
        throw 'No slide show is currently running.'
    }

    $view = $app.SlideShowWindows.Item(1).View

    switch ($Direction.ToLower()) {
        'next'     { $view.Next() }
        'previous' { $view.Previous() }
        'first'    { $view.First() }
        'last'     { $view.Last() }
        default {
            if ($Direction -match '^\d+$') {
                $view.GotoSlide([int]$Direction)
            }
            else {
                throw "Invalid direction '$Direction'. Use 'next', 'previous', 'first', 'last', or a slide number."
            }
        }
    }

    $result = [ordered]@{
        status       = 'navigated'
        direction    = $Direction
        currentSlide = $view.CurrentShowPosition
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointPresenterView
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointPresenterView {
    <#
    .SYNOPSIS
        Enable or disable Presenter View for the slide show.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER Enabled
        Whether Presenter View is enabled.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointPresenterView -Enabled $true -AsJson
    .EXAMPLE
        Set-PowerPointPresenterView -Enabled $false -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][bool]$Enabled,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation

    $triState = if ($Enabled) { -1 } else { 0 }   # msoTrue / msoFalse

    if (-not $PSCmdlet.ShouldProcess('SlideShowSettings.ShowPresenterView', "Set to $Enabled")) { return }

    $pres.SlideShowSettings.ShowPresenterView = $triState

    $result = [ordered]@{ status = 'updated'; presenterView = $Enabled }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
