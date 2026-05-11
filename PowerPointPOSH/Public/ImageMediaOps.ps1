# Public/ImageMediaOps.ps1 — Image, audio, video, and slide export operations

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointImage
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointImage {
    <#
    .SYNOPSIS
        Add an image to a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ImagePath
        Path to the image file. Must exist on disk.
    .PARAMETER Left
        Left position in points (default 100).
    .PARAMETER Top
        Top position in points (default 100).
    .PARAMETER Width
        Width in points. Omit to auto-size.
    .PARAMETER Height
        Height in points. Omit to auto-size.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointImage -SlideIndex 1 -ImagePath "C:\logo.png" -Left 50 -Top 50 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$ImagePath,
        [double]$Left   = 100,
        [double]$Top    = 100,
        [double]$Width  = -1,
        [double]$Height = -1,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $resolved = [System.IO.Path]::GetFullPath($ImagePath)
    if (-not (Test-Path -LiteralPath $resolved)) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.IO.FileNotFoundException]::new("Image file not found: $resolved"),
            'ImageNotFound',
            [System.Management.Automation.ErrorCategory]::ObjectNotFound, $resolved)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex", "Add image '$resolved'")) { return }

    # 0 = msoFalse (not linked), -1 = msoTrue (save with doc)
    if ($Width -le 0 -or $Height -le 0) {
        # Add with placeholder size, then let PowerPoint auto-size by locking aspect ratio
        $shape = $slide.Shapes.AddPicture($resolved, 0, -1, $Left, $Top, -1, -1)
        # PowerPoint may set tiny dims with -1; reset from the picture's natural size
        $shape.ScaleWidth(1, -1)   # msoTrue = relative to original
        $shape.ScaleHeight(1, -1)
    } else {
        $shape = $slide.Shapes.AddPicture($resolved, 0, -1, $Left, $Top, $Width, $Height)
    }

    $result = [ordered]@{
        status = 'added'
        name   = $shape.Name
        index  = $shape.ZOrderPosition
        width  = $shape.Width
        height = $shape.Height
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointAudio
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointAudio {
    <#
    .SYNOPSIS
        Add an audio clip to a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER AudioPath
        Path to the audio file. Must exist on disk.
    .PARAMETER Left
        Left position in points (default 100).
    .PARAMETER Top
        Top position in points (default 100).
    .PARAMETER PlayAcrossSlides
        Play audio across multiple slides.
    .PARAMETER HideIcon
        Hide the audio icon during playback.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointAudio -SlideIndex 1 -AudioPath "C:\music.mp3" -PlayAcrossSlides $true -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$AudioPath,
        [double]$Left = 100,
        [double]$Top  = 100,
        [bool]$PlayAcrossSlides,
        [bool]$HideIcon,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $resolved = [System.IO.Path]::GetFullPath($AudioPath)
    if (-not (Test-Path -LiteralPath $resolved)) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.IO.FileNotFoundException]::new("Audio file not found: $resolved"),
            'AudioNotFound',
            [System.Management.Automation.ErrorCategory]::ObjectNotFound, $resolved)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex", "Add audio '$resolved'")) { return }

    # 0 = msoFalse (not linked), -1 = msoTrue (save with doc)
    $shape = $slide.Shapes.AddMediaObject2($resolved, 0, -1, $Left, $Top)

    if ($PSBoundParameters.ContainsKey('PlayAcrossSlides') -and $PlayAcrossSlides) {
        $shape.AnimationSettings.PlaySettings.PlayOnEntry       = -1
        $shape.AnimationSettings.PlaySettings.PauseAnimation    = 0
        $shape.AnimationSettings.PlaySettings.StopAfterSlides   = 999
    }
    if ($PSBoundParameters.ContainsKey('HideIcon') -and $HideIcon) {
        $shape.AnimationSettings.PlaySettings.HideWhileNotPlaying = -1
    }

    $result = [ordered]@{
        status = 'added'
        name   = $shape.Name
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointVideo
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointVideo {
    <#
    .SYNOPSIS
        Add a video to a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER VideoPath
        Path to the video file. Must exist on disk.
    .PARAMETER Left
        Left position in points (default 100).
    .PARAMETER Top
        Top position in points (default 100).
    .PARAMETER Width
        Width in points (default -1, auto-size).
    .PARAMETER Height
        Height in points (default -1, auto-size).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointVideo -SlideIndex 1 -VideoPath "C:\demo.mp4" -Width 400 -Height 300 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$VideoPath,
        [double]$Left   = 100,
        [double]$Top    = 100,
        [double]$Width  = -1,
        [double]$Height = -1,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $resolved = [System.IO.Path]::GetFullPath($VideoPath)
    if (-not (Test-Path -LiteralPath $resolved)) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.IO.FileNotFoundException]::new("Video file not found: $resolved"),
            'VideoNotFound',
            [System.Management.Automation.ErrorCategory]::ObjectNotFound, $resolved)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex", "Add video '$resolved'")) { return }

    # 0 = msoFalse (not linked), -1 = msoTrue (save with doc)
    $shape = $slide.Shapes.AddMediaObject2($resolved, 0, -1, $Left, $Top, $Width, $Height)

    $result = [ordered]@{
        status = 'added'
        name   = $shape.Name
        width  = $shape.Width
        height = $shape.Height
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointMedia
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointMedia {
    <#
    .SYNOPSIS
        Get media information from a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the media shape.
    .PARAMETER ShapeIndex
        1-based index of the media shape.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointMedia -SlideIndex 1 -ShapeName "Video 1" -AsJson
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

    $mediaFormat = $shape.MediaFormat
    $mediaType   = try { [int]$shape.MediaType } catch { -1 }
    $fileName    = try { $mediaFormat.Filename }  catch { $null }
    $isEmbedded  = try { $mediaFormat.IsEmbedded } catch { $false }

    $result = [ordered]@{
        shapeName = $shape.Name
        mediaType = $mediaType
        fileName  = $fileName
        embedded  = [bool]$isEmbedded
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointMediaProperties
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointMediaProperties {
    <#
    .SYNOPSIS
        Set media playback properties (trim, fade, volume).
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the media shape.
    .PARAMETER ShapeIndex
        1-based index of the media shape.
    .PARAMETER StartTime
        Start trim time in milliseconds.
    .PARAMETER EndTime
        End trim time in milliseconds.
    .PARAMETER FadeInDuration
        Fade-in duration in milliseconds.
    .PARAMETER FadeOutDuration
        Fade-out duration in milliseconds.
    .PARAMETER Volume
        Playback volume (0.0 to 1.0).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointMediaProperties -SlideIndex 1 -ShapeName "Audio 1" -Volume 0.5 -FadeInDuration 2000 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [double]$StartTime,
        [double]$EndTime,
        [double]$FadeInDuration,
        [double]$FadeOutDuration,
        [double]$Volume,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Media '$target' on slide $SlideIndex", "Set media properties")) { return }

    $mf      = $shape.MediaFormat
    $changes = @()

    if ($PSBoundParameters.ContainsKey('StartTime'))       { $mf.TrimStart         = $StartTime;       $changes += 'startTime' }
    if ($PSBoundParameters.ContainsKey('EndTime'))         { $mf.TrimEnd           = $EndTime;         $changes += 'endTime' }
    if ($PSBoundParameters.ContainsKey('FadeInDuration'))  { $mf.FadeInDuration   = $FadeInDuration;  $changes += 'fadeInDuration' }
    if ($PSBoundParameters.ContainsKey('FadeOutDuration')) { $mf.FadeOutDuration  = $FadeOutDuration; $changes += 'fadeOutDuration' }
    if ($PSBoundParameters.ContainsKey('Volume'))          { $mf.Volume           = $Volume;          $changes += 'volume' }

    $result = [ordered]@{
        status  = 'updated'
        shape   = $shape.Name
        changed = $changes
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Remove-PowerPointMedia
# ══════════════════════════════════════════════════════════════════════════

function Remove-PowerPointMedia {
    <#
    .SYNOPSIS
        Remove a media shape from a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the media shape.
    .PARAMETER ShapeIndex
        1-based index of the media shape.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-PowerPointMedia -SlideIndex 1 -ShapeName "Audio 1" -AsJson
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

    $shapeName = $shape.Name
    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Media '$target' on slide $SlideIndex", "Remove media")) { return }

    $shape.Delete()

    $result = [ordered]@{
        status = 'removed'
        shape  = $shapeName
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Export-PowerPointSlideImage
# ══════════════════════════════════════════════════════════════════════════

function Export-PowerPointSlideImage {
    <#
    .SYNOPSIS
        Export a slide as an image file.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER FilePath
        Output file path for the exported image.
    .PARAMETER Format
        Image format: 'png', 'jpg', 'gif', or 'bmp' (default 'png').
    .PARAMETER Width
        Output image width in pixels.
    .PARAMETER Height
        Output image height in pixels.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Export-PowerPointSlideImage -SlideIndex 1 -FilePath "C:\slide1.png" -AsJson
    .EXAMPLE
        Export-PowerPointSlideImage -SlideIndex 2 -FilePath "C:\slide2.jpg" -Format jpg -Width 1920 -Height 1080 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$FilePath,
        [ValidateSet('png','jpg','gif','bmp')][string]$Format = 'png',
        [int]$Width,
        [int]$Height,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $resolvedFile = [System.IO.Path]::GetFullPath($FilePath)

    if ($PSBoundParameters.ContainsKey('Width') -and $PSBoundParameters.ContainsKey('Height')) {
        $slide.Export($resolvedFile, $Format, $Width, $Height)
    } elseif ($PSBoundParameters.ContainsKey('Width')) {
        $slide.Export($resolvedFile, $Format, $Width)
    } else {
        $slide.Export($resolvedFile, $Format)
    }

    $result = [ordered]@{
        status   = 'exported'
        filePath = $resolvedFile
        format   = $Format
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
