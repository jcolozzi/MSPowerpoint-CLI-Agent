# Public/SlideOps.ps1 — Slide operations: get, add, remove, copy, move, layout, notes, background, placeholders

function Get-PowerPointSlide {
    <#
    .SYNOPSIS
        Get information about one or all slides in the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based slide index. If omitted, returns all slides.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointSlide -SlideIndex 1 -AsJson
    .EXAMPLE
        Get-PowerPointSlide -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [int]$SlideIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointSlide'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $getSlideInfo = {
        param($slide)
        $layoutName = $null
        try { $layoutName = $slide.Layout } catch {}
        # Reverse-lookup layout friendly name
        $layoutKey = $layoutName
        foreach ($entry in $script:PPT_SLIDE_LAYOUT.GetEnumerator()) {
            if ($entry.Value -eq $layoutName) { $layoutKey = $entry.Key; break }
        }

        $hasNotes = $false
        try {
            $notesText = $slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            $hasNotes = -not [string]::IsNullOrWhiteSpace($notesText)
        } catch {}

        $transVal = $null
        try { $transVal = $slide.SlideShowTransition.EntryEffect } catch {}

        [ordered]@{
            index       = $slide.SlideIndex
            slideID     = $slide.SlideID
            layout      = $layoutKey
            name        = $slide.Name
            shapesCount = $slide.Shapes.Count
            hasNotes    = $hasNotes
            transition  = $transVal
        }
    }

    if ($PSBoundParameters.ContainsKey('SlideIndex')) {
        if ($SlideIndex -lt 1 -or $SlideIndex -gt $pres.Slides.Count) {
            $er = [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentOutOfRangeException]::new('SlideIndex', "Slide index $SlideIndex is out of range (1..$($pres.Slides.Count))."),
                'SlideIndexOutOfRange', [System.Management.Automation.ErrorCategory]::InvalidArgument, $SlideIndex)
            $PSCmdlet.ThrowTerminatingError($er)
        }
        $slide = $pres.Slides($SlideIndex)
        $result = & $getSlideInfo $slide
        $result.Insert(0, 'status', 'ok')
        return Format-PowerPointOutput -Data $result -AsJson:$AsJson
    }

    $slides = @()
    foreach ($slide in $pres.Slides) {
        $slides += & $getSlideInfo $slide
    }

    $result = [ordered]@{
        status     = 'ok'
        slideCount = $pres.Slides.Count
        slides     = $slides
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function New-PowerPointSlide {
    <#
    .SYNOPSIS
        Add a new slide to the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        Position at which to insert the slide (1-based). Default: end of presentation.
    .PARAMETER Layout
        Slide layout key (e.g., 'blank', 'title', 'twoColumnText'). Default: 'blank'.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-PowerPointSlide -Layout 'title' -AsJson
    .EXAMPLE
        New-PowerPointSlide -SlideIndex 2 -Layout 'blank'
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [int]$SlideIndex,
        [string]$Layout = 'blank',
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'New-PowerPointSlide'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $layoutValue = Resolve-EnumValue -Map $script:PPT_SLIDE_LAYOUT -Key $Layout

    if (-not $PSBoundParameters.ContainsKey('SlideIndex')) {
        $SlideIndex = $pres.Slides.Count + 1
    }

    $slide = $pres.Slides.Add($SlideIndex, $layoutValue)

    $result = [ordered]@{
        status     = 'ok'
        slideIndex = $slide.SlideIndex
        layout     = $Layout
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Remove-PowerPointSlide {
    <#
    .SYNOPSIS
        Remove a slide from the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide to remove.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-PowerPointSlide -SlideIndex 3 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Remove-PowerPointSlide'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($SlideIndex -lt 1 -or $SlideIndex -gt $pres.Slides.Count) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentOutOfRangeException]::new('SlideIndex', "Slide index $SlideIndex is out of range (1..$($pres.Slides.Count))."),
            'SlideIndexOutOfRange', [System.Management.Automation.ErrorCategory]::InvalidArgument, $SlideIndex)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    if ($PSCmdlet.ShouldProcess("Slide $SlideIndex", 'Remove')) {
        $pres.Slides($SlideIndex).Delete()
    }

    $result = [ordered]@{
        status       = 'ok'
        removedIndex = $SlideIndex
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Copy-PowerPointSlide {
    <#
    .SYNOPSIS
        Duplicate a slide within the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide to copy.
    .PARAMETER DestinationIndex
        1-based position to paste the copy. Default: end of presentation.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Copy-PowerPointSlide -SlideIndex 1 -DestinationIndex 3 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [int]$DestinationIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Copy-PowerPointSlide'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($SlideIndex -lt 1 -or $SlideIndex -gt $pres.Slides.Count) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentOutOfRangeException]::new('SlideIndex', "Slide index $SlideIndex is out of range (1..$($pres.Slides.Count))."),
            'SlideIndexOutOfRange', [System.Management.Automation.ErrorCategory]::InvalidArgument, $SlideIndex)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    if (-not $PSBoundParameters.ContainsKey('DestinationIndex')) {
        $DestinationIndex = $pres.Slides.Count + 1
    }

    $pres.Slides($SlideIndex).Copy()
    $pasted = $pres.Slides.Paste($DestinationIndex)

    $result = [ordered]@{
        status           = 'ok'
        sourceIndex      = $SlideIndex
        destinationIndex = $pasted.SlideIndex
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Move-PowerPointSlide {
    <#
    .SYNOPSIS
        Move a slide to a new position within the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide to move.
    .PARAMETER NewIndex
        1-based target position.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Move-PowerPointSlide -SlideIndex 3 -NewIndex 1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][int]$NewIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Move-PowerPointSlide'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($SlideIndex -lt 1 -or $SlideIndex -gt $pres.Slides.Count) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentOutOfRangeException]::new('SlideIndex', "Slide index $SlideIndex is out of range (1..$($pres.Slides.Count))."),
            'SlideIndexOutOfRange', [System.Management.Automation.ErrorCategory]::InvalidArgument, $SlideIndex)
        $PSCmdlet.ThrowTerminatingError($er)
    }
    if ($NewIndex -lt 1 -or $NewIndex -gt $pres.Slides.Count) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentOutOfRangeException]::new('NewIndex', "New index $NewIndex is out of range (1..$($pres.Slides.Count))."),
            'NewIndexOutOfRange', [System.Management.Automation.ErrorCategory]::InvalidArgument, $NewIndex)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $pres.Slides($SlideIndex).MoveTo($NewIndex)

    $result = [ordered]@{
        status = 'ok'
        from   = $SlideIndex
        to     = $NewIndex
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Set-PowerPointSlideLayout {
    <#
    .SYNOPSIS
        Change the layout of a slide.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide.
    .PARAMETER Layout
        Layout key (e.g., 'blank', 'title', 'twoColumnText').
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointSlideLayout -SlideIndex 1 -Layout 'titleOnly' -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$Layout,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Set-PowerPointSlideLayout'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($SlideIndex -lt 1 -or $SlideIndex -gt $pres.Slides.Count) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentOutOfRangeException]::new('SlideIndex', "Slide index $SlideIndex is out of range (1..$($pres.Slides.Count))."),
            'SlideIndexOutOfRange', [System.Management.Automation.ErrorCategory]::InvalidArgument, $SlideIndex)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $layoutValue = Resolve-EnumValue -Map $script:PPT_SLIDE_LAYOUT -Key $Layout

    if ($PSCmdlet.ShouldProcess("Slide $SlideIndex", "Set layout to '$Layout'")) {
        $pres.Slides($SlideIndex).Layout = $layoutValue
    }

    $result = [ordered]@{
        status     = 'ok'
        slideIndex = $SlideIndex
        layout     = $Layout
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Get-PowerPointSlideNotes {
    <#
    .SYNOPSIS
        Read the speaker notes from a slide.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointSlideNotes -SlideIndex 1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointSlideNotes'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($SlideIndex -lt 1 -or $SlideIndex -gt $pres.Slides.Count) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentOutOfRangeException]::new('SlideIndex', "Slide index $SlideIndex is out of range (1..$($pres.Slides.Count))."),
            'SlideIndexOutOfRange', [System.Management.Automation.ErrorCategory]::InvalidArgument, $SlideIndex)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $notes = ''
    try {
        $notes = $pres.Slides($SlideIndex).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
    } catch {
        $notes = ''
    }

    $result = [ordered]@{
        status     = 'ok'
        slideIndex = $SlideIndex
        notes      = $notes
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Set-PowerPointSlideNotes {
    <#
    .SYNOPSIS
        Set the speaker notes on a slide.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide.
    .PARAMETER Notes
        Text to set as speaker notes.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointSlideNotes -SlideIndex 1 -Notes "Remember to mention Q3 results" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$Notes,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Set-PowerPointSlideNotes'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($SlideIndex -lt 1 -or $SlideIndex -gt $pres.Slides.Count) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentOutOfRangeException]::new('SlideIndex', "Slide index $SlideIndex is out of range (1..$($pres.Slides.Count))."),
            'SlideIndexOutOfRange', [System.Management.Automation.ErrorCategory]::InvalidArgument, $SlideIndex)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    if ($PSCmdlet.ShouldProcess("Slide $SlideIndex", 'Set speaker notes')) {
        $pres.Slides($SlideIndex).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = $Notes
    }

    $result = [ordered]@{
        status     = 'ok'
        slideIndex = $SlideIndex
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Set-PowerPointSlideBackground {
    <#
    .SYNOPSIS
        Set the background of a slide to a solid color or an image.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide.
    .PARAMETER Color
        RGB color as "R,G,B" (e.g., "255,0,0") or hex (e.g., "#FF0000"). Mutually exclusive with -ImagePath.
    .PARAMETER ImagePath
        Path to an image file for the background. Mutually exclusive with -Color.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointSlideBackground -SlideIndex 1 -Color "0,0,128" -AsJson
    .EXAMPLE
        Set-PowerPointSlideBackground -SlideIndex 2 -ImagePath "C:\bg.jpg" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$Color,
        [string]$ImagePath,
        [switch]$AsJson
    )

    if ([string]::IsNullOrWhiteSpace($Color) -and [string]::IsNullOrWhiteSpace($ImagePath)) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentException]::new('Either -Color or -ImagePath must be specified.'),
            'MissingBackgroundParam', [System.Management.Automation.ErrorCategory]::InvalidArgument, $null)
        $PSCmdlet.ThrowTerminatingError($er)
    }
    if (-not [string]::IsNullOrWhiteSpace($Color) -and -not [string]::IsNullOrWhiteSpace($ImagePath)) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentException]::new('-Color and -ImagePath are mutually exclusive.'),
            'MutuallyExclusiveParams', [System.Management.Automation.ErrorCategory]::InvalidArgument, $null)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Set-PowerPointSlideBackground'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($SlideIndex -lt 1 -or $SlideIndex -gt $pres.Slides.Count) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentOutOfRangeException]::new('SlideIndex', "Slide index $SlideIndex is out of range (1..$($pres.Slides.Count))."),
            'SlideIndexOutOfRange', [System.Management.Automation.ErrorCategory]::InvalidArgument, $SlideIndex)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $slide = $pres.Slides($SlideIndex)
    $bgType = $null

    if (-not [string]::IsNullOrWhiteSpace($Color)) {
        # Parse color: hex "#RRGGBB" or RGB "R,G,B"
        $r = 0; $g = 0; $b = 0
        if ($Color -match '^#?([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})$') {
            $r = [Convert]::ToInt32($Matches[1], 16)
            $g = [Convert]::ToInt32($Matches[2], 16)
            $b = [Convert]::ToInt32($Matches[3], 16)
        }
        elseif ($Color -match '^\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*$') {
            $r = [int]$Matches[1]; $g = [int]$Matches[2]; $b = [int]$Matches[3]
        }
        else {
            $er = [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new("Invalid color format '$Color'. Use 'R,G,B' or '#RRGGBB'."),
                'InvalidColorFormat', [System.Management.Automation.ErrorCategory]::InvalidArgument, $Color)
            $PSCmdlet.ThrowTerminatingError($er)
        }

        # PowerPoint RGB = R + G*256 + B*65536
        $rgb = $r + ($g * 256) + ($b * 65536)

        if ($PSCmdlet.ShouldProcess("Slide $SlideIndex", "Set background color to $Color")) {
            $slide.FollowMasterBackground = 0  # msoFalse
            $slide.Background.Fill.Solid()
            $slide.Background.Fill.ForeColor.RGB = $rgb
        }
        $bgType = 'color'
    }
    else {
        $resolvedImage = [System.IO.Path]::GetFullPath($ImagePath)
        if (-not (Test-Path -LiteralPath $resolvedImage -PathType Leaf)) {
            $er = [System.Management.Automation.ErrorRecord]::new(
                [System.IO.FileNotFoundException]::new("Image file not found: $resolvedImage"),
                'ImageNotFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $resolvedImage)
            $PSCmdlet.ThrowTerminatingError($er)
        }

        if ($PSCmdlet.ShouldProcess("Slide $SlideIndex", "Set background image to $resolvedImage")) {
            $slide.FollowMasterBackground = 0  # msoFalse
            $slide.Background.Fill.UserPicture($resolvedImage)
        }
        $bgType = 'image'
    }

    $result = [ordered]@{
        status     = 'ok'
        slideIndex = $SlideIndex
        type       = $bgType
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Get-PowerPointSlidePlaceholders {
    <#
    .SYNOPSIS
        List all placeholders on a slide with their index, name, type, and text.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointSlidePlaceholders -SlideIndex 1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointSlidePlaceholders'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($SlideIndex -lt 1 -or $SlideIndex -gt $pres.Slides.Count) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentOutOfRangeException]::new('SlideIndex', "Slide index $SlideIndex is out of range (1..$($pres.Slides.Count))."),
            'SlideIndexOutOfRange', [System.Management.Automation.ErrorCategory]::InvalidArgument, $SlideIndex)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $placeholders = @()
    $slide = $pres.Slides($SlideIndex)

    foreach ($ph in $slide.Shapes.Placeholders) {
        $hasText = $false
        $text    = ''
        try {
            if ($ph.HasTextFrame) {
                $text    = $ph.TextFrame.TextRange.Text
                $hasText = -not [string]::IsNullOrWhiteSpace($text)
                # Truncate long text
                if ($text.Length -gt 200) { $text = $text.Substring(0, 200) + '...' }
            }
        } catch {}

        $typeName = $ph.PlaceholderFormat.Type
        # Reverse-lookup placeholder type friendly name
        foreach ($entry in $script:PPT_PLACEHOLDER.GetEnumerator()) {
            if ($entry.Value -eq $typeName) { $typeName = $entry.Key; break }
        }

        $placeholders += [ordered]@{
            index   = $ph.PlaceholderFormat.Type
            name    = $ph.Name
            type    = $typeName
            hasText = $hasText
            text    = $text
        }
    }

    $result = [ordered]@{
        status       = 'ok'
        slideIndex   = $SlideIndex
        placeholders = $placeholders
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
