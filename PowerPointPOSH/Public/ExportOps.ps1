# Public/ExportOps.ps1 — Export and convert presentations: PDF, images, video, HTML, single slide, format conversion

function Export-PowerPointToPdf {
    <#
    .SYNOPSIS
        Export the presentation to PDF.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER FilePath
        Destination PDF file path.
    .PARAMETER Quality
        PDF quality: 'standard' or 'minimum'. Default: 'standard'.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Export-PowerPointToPdf -FilePath "C:\output\deck.pdf" -AsJson
    .EXAMPLE
        Export-PowerPointToPdf -FilePath "C:\output\deck.pdf" -Quality 'minimum'
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$FilePath,
        [ValidateSet('standard','minimum')]
        [string]$Quality = 'standard',
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Export-PowerPointToPdf'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $FilePath = [System.IO.Path]::GetFullPath($FilePath)
    $dir = [System.IO.Path]::GetDirectoryName($FilePath)
    if (-not (Test-Path -LiteralPath $dir)) {
        $null = New-Item -Path $dir -ItemType Directory -Force
    }

    $formatType = Resolve-EnumValue -Map $script:PPT_FIXED_FORMAT -Key 'pdf'
    # PpFixedFormatIntent: 1 = ppFixedFormatIntentPrint, 2 = ppFixedFormatIntentScreen
    $intent = if ($Quality -eq 'minimum') { 2 } else { 1 }

    $pres.ExportAsFixedFormat($FilePath, $formatType, $intent)

    $result = [ordered]@{
        status   = 'ok'
        filePath = $FilePath
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Export-PowerPointToImages {
    <#
    .SYNOPSIS
        Export all slides as individual image files.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER OutputFolder
        Destination folder for exported images.
    .PARAMETER Format
        Image format: 'png', 'jpg', 'gif', or 'bmp'. Default: 'png'.
    .PARAMETER Width
        Image width in pixels. Optional.
    .PARAMETER Height
        Image height in pixels. Optional.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Export-PowerPointToImages -OutputFolder "C:\slides" -Format 'png' -AsJson
    .EXAMPLE
        Export-PowerPointToImages -OutputFolder "C:\slides" -Format 'jpg' -Width 1920 -Height 1080
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$OutputFolder,
        [ValidateSet('png','jpg','gif','bmp')]
        [string]$Format = 'png',
        [int]$Width,
        [int]$Height,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Export-PowerPointToImages'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $OutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)
    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        $null = New-Item -Path $OutputFolder -ItemType Directory -Force
    }

    $count = 0
    foreach ($slide in $pres.Slides) {
        $fileName = Join-Path $OutputFolder "Slide$($slide.SlideIndex).$Format"
        if ($PSBoundParameters.ContainsKey('Width') -and $PSBoundParameters.ContainsKey('Height')) {
            $slide.Export($fileName, $Format, $Width, $Height)
        } elseif ($PSBoundParameters.ContainsKey('Width')) {
            $slide.Export($fileName, $Format, $Width)
        } else {
            $slide.Export($fileName, $Format)
        }
        $count++
    }

    $result = [ordered]@{
        status       = 'ok'
        outputFolder = $OutputFolder
        slideCount   = $count
        format       = $Format
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Export-PowerPointToVideo {
    <#
    .SYNOPSIS
        Export the presentation as a video file (MP4 or WMV).
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER FilePath
        Destination video file path.
    .PARAMETER Format
        Video format: 'mp4' or 'wmv'. Default: 'mp4'.
    .PARAMETER DefaultSlideDuration
        Duration per slide in seconds. Default: 5.
    .PARAMETER Resolution
        Video resolution: 480, 720, or 1080. Default: 720.
    .PARAMETER UseTimingsAndNarrations
        Use slide timings and narrations. Default: $true.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Export-PowerPointToVideo -FilePath "C:\output\deck.mp4" -AsJson
    .EXAMPLE
        Export-PowerPointToVideo -FilePath "C:\output\deck.wmv" -Format 'wmv' -Resolution 1080 -DefaultSlideDuration 3
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$FilePath,
        [ValidateSet('mp4','wmv')]
        [string]$Format = 'mp4',
        [int]$DefaultSlideDuration = 5,
        [ValidateSet(480, 720, 1080)]
        [int]$Resolution = 720,
        [bool]$UseTimingsAndNarrations = $true,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Export-PowerPointToVideo'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $FilePath = [System.IO.Path]::GetFullPath($FilePath)
    $dir = [System.IO.Path]::GetDirectoryName($FilePath)
    if (-not (Test-Path -LiteralPath $dir)) {
        $null = New-Item -Path $dir -ItemType Directory -Force
    }

    $useTimings = if ($UseTimingsAndNarrations) { -1 } else { 0 }  # msoTrue / msoFalse

    $pres.CreateVideo($FilePath, $useTimings, $DefaultSlideDuration, $Resolution)

    # CreateVideo is async — poll until complete
    # ppMediaTaskStatusInProgress = 1, ppMediaTaskStatusQueued = 2, ppMediaTaskStatusDone = 3
    $maxWait = 600  # 10 minutes max
    $elapsed = 0
    while ($pres.CreateVideoStatus -eq 1 -or $pres.CreateVideoStatus -eq 2) {
        Start-Sleep -Seconds 1
        $elapsed++
        if ($elapsed -ge $maxWait) {
            throw "Video creation timed out after $maxWait seconds."
        }
    }

    if ($pres.CreateVideoStatus -ne 3) {
        throw "Video creation failed with status: $($pres.CreateVideoStatus)"
    }

    $result = [ordered]@{
        status   = 'ok'
        filePath = $FilePath
        format   = $Format
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Export-PowerPointToHtml {
    <#
    .SYNOPSIS
        Export the presentation to HTML format.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER FilePath
        Destination HTML file path.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Export-PowerPointToHtml -FilePath "C:\output\deck.html" -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$FilePath,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Export-PowerPointToHtml'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $FilePath = [System.IO.Path]::GetFullPath($FilePath)
    $dir = [System.IO.Path]::GetDirectoryName($FilePath)
    if (-not (Test-Path -LiteralPath $dir)) {
        $null = New-Item -Path $dir -ItemType Directory -Force
    }

    # ppSaveAsHTML = 13 (legacy format, may not be available in all versions)
    try {
        $pres.SaveAs($FilePath, 13)
    } catch {
        throw "HTML export failed. This format (ppSaveAsHTML) may not be supported in your version of PowerPoint. Error: $_"
    }

    $result = [ordered]@{
        status   = 'ok'
        filePath = $FilePath
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Export-PowerPointSlide {
    <#
    .SYNOPSIS
        Export a single slide as an image file.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide to export.
    .PARAMETER FilePath
        Destination image file path.
    .PARAMETER Format
        Image format: 'png' or 'jpg'. Default: 'png'.
    .PARAMETER Width
        Image width in pixels. Optional.
    .PARAMETER Height
        Image height in pixels. Optional.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Export-PowerPointSlide -SlideIndex 1 -FilePath "C:\slide1.png" -AsJson
    .EXAMPLE
        Export-PowerPointSlide -SlideIndex 3 -FilePath "C:\slide3.jpg" -Format 'jpg' -Width 1920 -Height 1080
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$FilePath,
        [ValidateSet('png','jpg')]
        [string]$Format = 'png',
        [int]$Width,
        [int]$Height,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Export-PowerPointSlide'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $FilePath = [System.IO.Path]::GetFullPath($FilePath)
    $dir = [System.IO.Path]::GetDirectoryName($FilePath)
    if (-not (Test-Path -LiteralPath $dir)) {
        $null = New-Item -Path $dir -ItemType Directory -Force
    }

    $slide = $pres.Slides.Item($SlideIndex)

    if ($PSBoundParameters.ContainsKey('Width') -and $PSBoundParameters.ContainsKey('Height')) {
        $slide.Export($FilePath, $Format, $Width, $Height)
    } elseif ($PSBoundParameters.ContainsKey('Width')) {
        $slide.Export($FilePath, $Format, $Width)
    } else {
        $slide.Export($FilePath, $Format)
    }

    $result = [ordered]@{
        status   = 'ok'
        filePath = $FilePath
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Convert-PowerPointFormat {
    <#
    .SYNOPSIS
        Convert a presentation to a different file format.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER DestinationPath
        Destination file path for the converted presentation.
    .PARAMETER Format
        Target format key (e.g., 'pptx', 'pptm', 'ppt', 'pdf', 'odp'). Resolved via PPT_FILE_FORMAT.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Convert-PowerPointFormat -DestinationPath "C:\deck.pptm" -Format 'pptm' -AsJson
    .EXAMPLE
        Convert-PowerPointFormat -DestinationPath "C:\deck.odp" -Format 'odp'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$DestinationPath,
        [Parameter(Mandatory)][string]$Format,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Convert-PowerPointFormat'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $DestinationPath = [System.IO.Path]::GetFullPath($DestinationPath)
    $formatValue = Resolve-EnumValue -Map $script:PPT_FILE_FORMAT -Key $Format

    if ($PSCmdlet.ShouldProcess($DestinationPath, "Convert presentation to '$Format' format")) {
        $pres.SaveAs($DestinationPath, $formatValue)
    }

    $result = [ordered]@{
        status      = 'ok'
        source      = $pres.FullName
        destination = $DestinationPath
        format      = $Format
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
