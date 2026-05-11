# Public/MetadataOps.ps1 — Document properties, comments, and tags

function Get-PowerPointDocumentProperty {
    <#
    .SYNOPSIS
        Get built-in document properties of the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER PropertyName
        Specific property name to retrieve. If omitted, returns all built-in properties.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointDocumentProperty -AsJson
    .EXAMPLE
        Get-PowerPointDocumentProperty -PropertyName 'Author' -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [string]$PropertyName,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointDocumentProperty'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $builtIn = $pres.BuiltInDocumentProperties

    if ($PSBoundParameters.ContainsKey('PropertyName')) {
        $val = $null
        try { $val = $builtIn.Item($PropertyName).Value } catch {}
        $result = [ordered]@{
            status   = 'ok'
            property = $PropertyName
            value    = $val
        }
        return Format-PowerPointOutput -Data $result -AsJson:$AsJson
    }

    # Retrieve all known built-in properties
    $knownProps = @(
        'Title','Subject','Author','Keywords','Comments','Category',
        'Last Author','Creation Date','Last Save Time','Revision Number',
        'Total Editing Time','Number of Slides','Number of Words',
        'Company','Manager','Template'
    )
    $props = [ordered]@{}
    foreach ($name in $knownProps) {
        try {
            $props[$name] = $builtIn.Item($name).Value
        } catch {
            $props[$name] = $null
        }
    }

    $result = [ordered]@{
        status     = 'ok'
        properties = $props
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Set-PowerPointDocumentProperty {
    <#
    .SYNOPSIS
        Set a built-in document property.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER PropertyName
        Name of the built-in property to set (e.g., 'Title', 'Author').
    .PARAMETER Value
        Value to assign to the property.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointDocumentProperty -PropertyName 'Title' -Value 'My Presentation' -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$PropertyName,
        [Parameter(Mandatory)][string]$Value,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Set-PowerPointDocumentProperty'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($PSCmdlet.ShouldProcess("$PropertyName = '$Value'", 'Set document property')) {
        $pres.BuiltInDocumentProperties.Item($PropertyName).Value = $Value
    }

    $result = [ordered]@{
        status   = 'ok'
        property = $PropertyName
        value    = $Value
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Get-PowerPointComment {
    <#
    .SYNOPSIS
        Get comments from one or all slides.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based slide index. If omitted, returns comments from all slides.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointComment -AsJson
    .EXAMPLE
        Get-PowerPointComment -SlideIndex 1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [int]$SlideIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointComment'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $getComments = {
        param($slide)
        $list = @()
        foreach ($c in $slide.Comments) {
            $dt = $null
            try { $dt = $c.DateTime.ToString('o') } catch {}
            $list += [ordered]@{
                slideIndex     = $slide.SlideIndex
                commentIndex   = $c.Index
                author         = $c.Author
                authorInitials = $c.AuthorInitials
                text           = $c.Text
                dateTime       = $dt
                left           = $c.Left
                top            = $c.Top
            }
        }
        $list
    }

    $comments = @()
    if ($PSBoundParameters.ContainsKey('SlideIndex')) {
        $slide = $pres.Slides.Item($SlideIndex)
        $comments = & $getComments $slide
    } else {
        foreach ($slide in $pres.Slides) {
            $comments += & $getComments $slide
        }
    }

    $result = [ordered]@{
        status       = 'ok'
        commentCount = $comments.Count
        comments     = $comments
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Add-PowerPointComment {
    <#
    .SYNOPSIS
        Add a comment to a slide.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the target slide.
    .PARAMETER Text
        Comment text.
    .PARAMETER Left
        Horizontal position in points. Default: 10.
    .PARAMETER Top
        Vertical position in points. Default: 10.
    .PARAMETER Author
        Comment author name. Default: current user.
    .PARAMETER AuthorInitials
        Comment author initials. Default: derived from Author.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointComment -SlideIndex 1 -Text 'Needs revision' -AsJson
    .EXAMPLE
        Add-PowerPointComment -SlideIndex 2 -Text 'Good chart' -Author 'JD' -AuthorInitials 'JD' -Left 100 -Top 50
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$Text,
        [double]$Left = 10,
        [double]$Top = 10,
        [string]$Author,
        [string]$AuthorInitials,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Add-PowerPointComment'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if (-not $PSBoundParameters.ContainsKey('Author')) {
        $Author = $env:USERNAME
    }
    if (-not $PSBoundParameters.ContainsKey('AuthorInitials')) {
        # Derive initials from author name
        $AuthorInitials = ($Author -split '\s+' | ForEach-Object { $_[0] }) -join ''
        if ([string]::IsNullOrWhiteSpace($AuthorInitials)) { $AuthorInitials = $Author }
    }

    $slide = $pres.Slides.Item($SlideIndex)

    if ($PSCmdlet.ShouldProcess("Slide $SlideIndex", "Add comment '$Text'")) {
        $null = $slide.Comments.Add($Left, $Top, $Author, $AuthorInitials, $Text)
    }

    $result = [ordered]@{
        status = 'ok'
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Remove-PowerPointComment {
    <#
    .SYNOPSIS
        Remove a comment from a slide.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based index of the slide containing the comment.
    .PARAMETER CommentIndex
        1-based index of the comment on the slide.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-PowerPointComment -SlideIndex 1 -CommentIndex 1 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][int]$CommentIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Remove-PowerPointComment'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $slide = $pres.Slides.Item($SlideIndex)

    if ($PSCmdlet.ShouldProcess("Slide $SlideIndex, Comment $CommentIndex", 'Remove comment')) {
        $slide.Comments.Item($CommentIndex).Delete()
    }

    $result = [ordered]@{
        status = 'ok'
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Get-PowerPointTag {
    <#
    .SYNOPSIS
        Get tags from a slide or the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SlideIndex
        1-based slide index. If omitted, returns presentation-level tags.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointTag -AsJson
    .EXAMPLE
        Get-PowerPointTag -SlideIndex 1 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [int]$SlideIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointTag'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $obj = if ($PSBoundParameters.ContainsKey('SlideIndex')) {
        $pres.Slides.Item($SlideIndex)
    } else {
        $pres
    }

    $tags = $obj.Tags
    $tagList = @()
    for ($i = 1; $i -le $tags.Count; $i++) {
        $tagList += [ordered]@{
            name  = $tags.Name($i)
            value = $tags.Value($i)
        }
    }

    $result = [ordered]@{
        status   = 'ok'
        tagCount = $tagList.Count
        tags     = $tagList
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Set-PowerPointTag {
    <#
    .SYNOPSIS
        Set a tag on a slide or the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER Name
        Tag name.
    .PARAMETER Value
        Tag value.
    .PARAMETER SlideIndex
        1-based slide index. If omitted, sets a presentation-level tag.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointTag -Name 'Status' -Value 'Draft' -AsJson
    .EXAMPLE
        Set-PowerPointTag -SlideIndex 1 -Name 'ReviewedBy' -Value 'JD' -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Value,
        [int]$SlideIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Set-PowerPointTag'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $obj = if ($PSBoundParameters.ContainsKey('SlideIndex')) {
        $pres.Slides.Item($SlideIndex)
    } else {
        $pres
    }

    $target = if ($PSBoundParameters.ContainsKey('SlideIndex')) { "Slide $SlideIndex" } else { 'Presentation' }

    if ($PSCmdlet.ShouldProcess($target, "Set tag '$Name' = '$Value'")) {
        $obj.Tags.Add($Name, $Value)
    }

    $result = [ordered]@{
        status = 'ok'
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
