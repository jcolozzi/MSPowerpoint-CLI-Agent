# Public/TextOps.ps1 — Text content and formatting operations

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointText
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointText {
    <#
    .SYNOPSIS
        Read text content and font info from a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointText -SlideIndex 1 -ShapeName "Title 1" -AsJson
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

    if ($shape.HasTextFrame -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not have a text frame."),
            'NoTextFrame',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $textRange = $shape.TextFrame.TextRange
    $font      = $textRange.Font

    $result = [ordered]@{
        text      = $textRange.Text
        fontName  = $font.Name
        fontSize  = $font.Size
        bold      = $font.Bold -eq -1
        italic    = $font.Italic -eq -1
        color     = $font.Color.RGB
        shapeName = $shape.Name
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointText
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointText {
    <#
    .SYNOPSIS
        Set the text content of a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER Text
        Text content to set.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointText -SlideIndex 1 -ShapeName "Title 1" -Text "New Title" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][string]$Text,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Set text")) { return }

    $shape.TextFrame.TextRange.Text = $Text

    $result = [ordered]@{
        status = 'updated'
        shape  = $shape.Name
        text   = $Text
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Format-PowerPointTextFont
# ══════════════════════════════════════════════════════════════════════════

function Format-PowerPointTextFont {
    <#
    .SYNOPSIS
        Format the font of text in a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER FontName
        Font family name (e.g. 'Calibri', 'Arial').
    .PARAMETER FontSize
        Font size in points.
    .PARAMETER Bold
        Set bold formatting.
    .PARAMETER Italic
        Set italic formatting.
    .PARAMETER Underline
        Set underline formatting.
    .PARAMETER Color
        Font color as 'R,G,B' string (e.g. '255,0,0').
    .PARAMETER StrikeThrough
        Set strikethrough formatting.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Format-PowerPointTextFont -SlideIndex 1 -ShapeName "Title 1" -FontName "Arial" -FontSize 24 -Bold $true -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [string]$FontName,
        [double]$FontSize,
        [bool]$Bold,
        [bool]$Italic,
        [bool]$Underline,
        [string]$Color,
        [bool]$StrikeThrough,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Format font")) { return }

    $font    = $shape.TextFrame.TextRange.Font
    $changed = @()

    if ($PSBoundParameters.ContainsKey('FontName'))      { $font.Name          = $FontName;                                        $changed += 'FontName' }
    if ($PSBoundParameters.ContainsKey('FontSize'))       { $font.Size          = $FontSize;                                        $changed += 'FontSize' }
    if ($PSBoundParameters.ContainsKey('Bold'))           { $font.Bold          = if ($Bold)          { -1 } else { 0 };            $changed += 'Bold' }
    if ($PSBoundParameters.ContainsKey('Italic'))         { $font.Italic        = if ($Italic)        { -1 } else { 0 };            $changed += 'Italic' }
    if ($PSBoundParameters.ContainsKey('Underline'))      { $font.Underline     = if ($Underline)     { -1 } else { 0 };            $changed += 'Underline' }
    if ($PSBoundParameters.ContainsKey('StrikeThrough'))  { $font.StrikeThrough = if ($StrikeThrough) { -1 } else { 0 };            $changed += 'StrikeThrough' }
    if ($PSBoundParameters.ContainsKey('Color'))          { $font.Color.RGB     = ConvertFrom-RGBString $Color;                     $changed += 'Color' }

    $result = [ordered]@{
        status  = 'formatted'
        shape   = $shape.Name
        changed = $changed
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Format-PowerPointTextParagraph
# ══════════════════════════════════════════════════════════════════════════

function Format-PowerPointTextParagraph {
    <#
    .SYNOPSIS
        Format paragraph properties of text in a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER Alignment
        Paragraph alignment key (e.g. 'left', 'center', 'right'). Resolved via PPT_ALIGN map.
    .PARAMETER SpaceBefore
        Space before paragraph in points.
    .PARAMETER SpaceAfter
        Space after paragraph in points.
    .PARAMETER SpaceWithin
        Line spacing within paragraph (multiplier, e.g. 1.5).
    .PARAMETER IndentLevel
        Indent level (1-5).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Format-PowerPointTextParagraph -SlideIndex 1 -ShapeName "Body" -Alignment center -SpaceBefore 6 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [string]$Alignment,
        [double]$SpaceBefore,
        [double]$SpaceAfter,
        [double]$SpaceWithin,
        [int]$IndentLevel,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Format paragraph")) { return }

    $textRange = $shape.TextFrame.TextRange
    $paraFmt   = $textRange.ParagraphFormat
    $changed   = @()

    if ($PSBoundParameters.ContainsKey('Alignment'))   { $paraFmt.Alignment   = Resolve-EnumValue -Map $script:PPT_ALIGN -Key $Alignment; $changed += 'Alignment' }
    if ($PSBoundParameters.ContainsKey('SpaceBefore')) { $paraFmt.SpaceBefore = $SpaceBefore; $changed += 'SpaceBefore' }
    if ($PSBoundParameters.ContainsKey('SpaceAfter'))  { $paraFmt.SpaceAfter  = $SpaceAfter;  $changed += 'SpaceAfter' }
    if ($PSBoundParameters.ContainsKey('SpaceWithin')) { $paraFmt.SpaceWithin = $SpaceWithin; $changed += 'SpaceWithin' }
    if ($PSBoundParameters.ContainsKey('IndentLevel')) { $textRange.IndentLevel = $IndentLevel; $changed += 'IndentLevel' }

    $result = [ordered]@{
        status  = 'formatted'
        shape   = $shape.Name
        changed = $changed
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointBullet
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointBullet {
    <#
    .SYNOPSIS
        Configure bullet formatting for text in a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the shape.
    .PARAMETER ShapeIndex
        1-based index of the shape.
    .PARAMETER Type
        Bullet type: 'numbered', 'unnumbered', or 'none'.
    .PARAMETER Character
        Custom bullet character (e.g. '•', '→'). Used with 'unnumbered' type.
    .PARAMETER StartValue
        Starting number for numbered bullets.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointBullet -SlideIndex 1 -ShapeName "Body" -Type numbered -StartValue 1 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][ValidateSet('numbered','unnumbered','none')][string]$Type,
        [string]$Character,
        [int]$StartValue,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Shape '$target' on slide $SlideIndex", "Set bullet type '$Type'")) { return }

    $bullet = $shape.TextFrame.TextRange.ParagraphFormat.Bullet

    # ppBulletNone=0, ppBulletUnnumbered=1, ppBulletNumbered=2
    $bulletType = switch ($Type.ToLower()) {
        'none'       { 0 }
        'unnumbered' { 1 }
        'numbered'   { 2 }
    }
    $bullet.Type = $bulletType

    if ($Character -and $Type -eq 'unnumbered') {
        $bullet.Character = [int][char]$Character[0]
    }
    if ($PSBoundParameters.ContainsKey('StartValue') -and $Type -eq 'numbered') {
        $bullet.StartValue = $StartValue
    }

    $result = [ordered]@{
        status     = 'configured'
        shape      = $shape.Name
        bulletType = $Type
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Find-PowerPointText
# ══════════════════════════════════════════════════════════════════════════

function Find-PowerPointText {
    <#
    .SYNOPSIS
        Search all slides and shapes for text content.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SearchText
        Text string to search for.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Find-PowerPointText -SearchText "revenue" -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$SearchText,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app  = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres = $app.ActivePresentation

    $result = @()
    foreach ($slide in $pres.Slides) {
        foreach ($s in $slide.Shapes) {
            $hasText = $false
            try { if ($s.HasTextFrame -eq -1 -and $s.TextFrame.HasText -eq -1) { $hasText = $true } } catch {}
            if (-not $hasText) { continue }

            $text = $s.TextFrame.TextRange.Text
            if ($text -like "*$SearchText*") {
                $matchCount  = ([regex]::Matches($text, [regex]::Escape($SearchText), 'IgnoreCase')).Count
                $displayText = if ($text.Length -gt 100) { $text.Substring(0, 100) + '...' } else { $text }
                $result += [ordered]@{
                    slideIndex = $slide.SlideIndex
                    shapeName  = $s.Name
                    text       = $displayText
                    matchCount = $matchCount
                }
            }
        }
    }

    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointTextReplace
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointTextReplace {
    <#
    .SYNOPSIS
        Find and replace text across all slides and shapes.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER FindText
        Text string to find.
    .PARAMETER ReplaceText
        Replacement text string.
    .PARAMETER MatchCase
        Perform case-sensitive matching.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointTextReplace -FindText "2024" -ReplaceText "2025" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$FindText,
        [Parameter(Mandatory)][string]$ReplaceText,
        [bool]$MatchCase,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app  = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres = $app.ActivePresentation

    if (-not $PSCmdlet.ShouldProcess("All slides", "Replace '$FindText' with '$ReplaceText'")) { return }

    $comparison      = if ($MatchCase) { [System.StringComparison]::Ordinal } else { [System.StringComparison]::OrdinalIgnoreCase }
    $replacementCount = 0

    foreach ($slide in $pres.Slides) {
        foreach ($s in $slide.Shapes) {
            $hasText = $false
            try { if ($s.HasTextFrame -eq -1 -and $s.TextFrame.HasText -eq -1) { $hasText = $true } } catch {}
            if (-not $hasText) { continue }

            $tr      = $s.TextFrame.TextRange
            $current = $tr.Text
            if ($current.IndexOf($FindText, $comparison) -ge 0) {
                $pattern = [regex]::Escape($FindText)
                $options = if ($MatchCase) { [System.Text.RegularExpressions.RegexOptions]::None } else { [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
                $count   = ([regex]::Matches($current, $pattern, $options)).Count
                $tr.Text = [regex]::Replace($current, $pattern, $ReplaceText, $options)
                $replacementCount += $count
            }
        }
    }

    $result = [ordered]@{
        status           = 'replaced'
        findText         = $FindText
        replaceText      = $ReplaceText
        replacementCount = $replacementCount
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointTextAll
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointTextAll {
    <#
    .SYNOPSIS
        Extract all text from all slides in the presentation.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointTextAll -PresentationPath "C:\deck.pptx" -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app  = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres = $app.ActivePresentation

    $result = @()
    foreach ($slide in $pres.Slides) {
        foreach ($s in $slide.Shapes) {
            $hasText = $false
            try { if ($s.HasTextFrame -eq -1 -and $s.TextFrame.HasText -eq -1) { $hasText = $true } } catch {}
            if (-not $hasText) { continue }

            $result += [ordered]@{
                slideIndex = $slide.SlideIndex
                shapeName  = $s.Name
                text       = $s.TextFrame.TextRange.Text
            }
        }
    }

    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
