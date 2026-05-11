# Public/PrintOps.ps1 — Page setup and printing operations

function Set-PowerPointPageSetup {
    <#
    .SYNOPSIS
        Modify page setup properties of the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideWidth
        Slide width in points.
    .PARAMETER SlideHeight
        Slide height in points.
    .PARAMETER SlideSize
        Slide size key (e.g., 'onScreen', 'widescreen', 'a4Paper'). Resolved via PPT_SLIDE_SIZE.
    .PARAMETER FirstSlideNumber
        Starting slide number.
    .PARAMETER Orientation
        Slide orientation: 'landscape' or 'portrait'.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointPageSetup -SlideSize 'widescreen' -AsJson
    .EXAMPLE
        Set-PowerPointPageSetup -SlideWidth 720 -SlideHeight 540 -Orientation 'landscape'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [double]$SlideWidth,
        [double]$SlideHeight,
        [string]$SlideSize,
        [int]$FirstSlideNumber,
        [ValidateSet('landscape','portrait')]
        [string]$Orientation,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Set-PowerPointPageSetup'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($PSCmdlet.ShouldProcess($pres.Name, 'Modify page setup')) {
        $ps = $pres.PageSetup

        if ($PSBoundParameters.ContainsKey('SlideSize')) {
            $sizeValue = Resolve-EnumValue -Map $script:PPT_SLIDE_SIZE -Key $SlideSize
            $ps.SlideSize = $sizeValue
        }

        if ($PSBoundParameters.ContainsKey('SlideWidth'))  { $ps.SlideWidth  = $SlideWidth }
        if ($PSBoundParameters.ContainsKey('SlideHeight')) { $ps.SlideHeight = $SlideHeight }

        if ($PSBoundParameters.ContainsKey('FirstSlideNumber')) {
            $ps.FirstSlideNumber = $FirstSlideNumber
        }

        if ($PSBoundParameters.ContainsKey('Orientation')) {
            # msoOrientationPortrait = 1, msoOrientationLandscape = 2
            $ps.SlideOrientation = if ($Orientation -eq 'landscape') { 2 } else { 1 }
        }
    }

    # Return current settings
    $ps = $pres.PageSetup
    $orientName = if ($ps.SlideOrientation -eq 2) { 'landscape' } else { 'portrait' }

    $result = [ordered]@{
        status           = 'ok'
        slideWidth       = $ps.SlideWidth
        slideHeight      = $ps.SlideHeight
        slideSize        = $ps.SlideSize
        firstSlideNumber = $ps.FirstSlideNumber
        orientation      = $orientName
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Get-PowerPointPageSetup {
    <#
    .SYNOPSIS
        Get page setup properties of the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointPageSetup -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointPageSetup'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $ps = $pres.PageSetup
    $orientName      = if ($ps.SlideOrientation -eq 2) { 'landscape' } else { 'portrait' }
    $notesOrientName = try {
        if ($ps.NotesOrientation -eq 2) { 'landscape' } else { 'portrait' }
    } catch { 'unknown' }

    $result = [ordered]@{
        status           = 'ok'
        slideWidth       = $ps.SlideWidth
        slideHeight      = $ps.SlideHeight
        slideSize        = $ps.SlideSize
        firstSlideNumber = $ps.FirstSlideNumber
        orientation      = $orientName
        notesOrientation = $notesOrientName
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Invoke-PowerPointPrint {
    <#
    .SYNOPSIS
        Print the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER PrintRange
        Print range: 'all', 'current', 'selection', or 'range'. Default: 'all'.
    .PARAMETER From
        Starting slide number (used with -PrintRange 'range').
    .PARAMETER To
        Ending slide number (used with -PrintRange 'range').
    .PARAMETER Copies
        Number of copies. Default: 1.
    .PARAMETER OutputType
        Print output type key (e.g., 'slides', 'notesPages', 'outline'). Resolved via PPT_PRINT_OUTPUT.
    .PARAMETER Collate
        Collate copies.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Invoke-PowerPointPrint -AsJson
    .EXAMPLE
        Invoke-PowerPointPrint -PrintRange 'range' -From 1 -To 5 -Copies 2 -OutputType 'notesPages'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [ValidateSet('all','current','selection','range')]
        [string]$PrintRange = 'all',
        [int]$From,
        [int]$To,
        [int]$Copies = 1,
        [string]$OutputType,
        [bool]$Collate,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Invoke-PowerPointPrint'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($PSCmdlet.ShouldProcess($pres.Name, "Print presentation ($PrintRange)")) {

        if ($PSBoundParameters.ContainsKey('OutputType')) {
            $outputValue = Resolve-EnumValue -Map $script:PPT_PRINT_OUTPUT -Key $OutputType
            $app.ActivePresentation.PrintOptions.OutputType = $outputValue
        }

        if ($PSBoundParameters.ContainsKey('Collate')) {
            $pres.PrintOptions.Collate = if ($Collate) { -1 } else { 0 }
        }

        switch ($PrintRange) {
            'range' {
                if (-not $PSBoundParameters.ContainsKey('From') -or -not $PSBoundParameters.ContainsKey('To')) {
                    throw "Parameters -From and -To are required when -PrintRange is 'range'."
                }
                $pres.PrintOut($From, $To, '', $Copies)
            }
            default {
                $pres.PrintOut(0, 0, '', $Copies)
            }
        }
    }

    $result = [ordered]@{
        status = 'ok'
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
