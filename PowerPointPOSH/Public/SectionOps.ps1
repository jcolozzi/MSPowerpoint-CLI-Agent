# Public/SectionOps.ps1 — Section operations: get, add, remove, rename, move

function Get-PowerPointSection {
    <#
    .SYNOPSIS
        Get all sections in the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointSection -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointSection'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $sp = $pres.SectionProperties
    $sections = @()
    for ($i = 1; $i -le $sp.Count; $i++) {
        $sections += [ordered]@{
            index       = $i
            name        = $sp.Name($i)
            firstSlide  = $sp.FirstSlide($i)
            slidesCount = $sp.SlidesCount($i)
        }
    }

    $result = [ordered]@{
        status       = 'ok'
        sectionCount = $sections.Count
        sections     = $sections
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function New-PowerPointSection {
    <#
    .SYNOPSIS
        Add a new section to the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER Name
        Name for the new section.
    .PARAMETER FirstSlideIndex
        1-based index of the first slide in the new section.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        New-PowerPointSection -Name 'Introduction' -FirstSlideIndex 1 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][int]$FirstSlideIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'New-PowerPointSection'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    if ($PSCmdlet.ShouldProcess($pres.Name, "Add section '$Name' at slide $FirstSlideIndex")) {
        $sectionIndex = $pres.SectionProperties.AddSection($FirstSlideIndex, $Name)
    }

    $result = [ordered]@{
        status       = 'ok'
        sectionIndex = $sectionIndex
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Remove-PowerPointSection {
    <#
    .SYNOPSIS
        Remove a section from the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SectionIndex
        1-based index of the section to remove.
    .PARAMETER DeleteSlides
        Also delete the slides in the section. Default: $false.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-PowerPointSection -SectionIndex 2 -AsJson
    .EXAMPLE
        Remove-PowerPointSection -SectionIndex 3 -DeleteSlides $true -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SectionIndex,
        [bool]$DeleteSlides = $false,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Remove-PowerPointSection'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $sectionName = try { $pres.SectionProperties.Name($SectionIndex) } catch { "Section $SectionIndex" }

    if ($PSCmdlet.ShouldProcess($sectionName, "Remove section (DeleteSlides=$DeleteSlides)")) {
        $pres.SectionProperties.Delete($SectionIndex, $DeleteSlides)
    }

    $result = [ordered]@{
        status = 'ok'
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Rename-PowerPointSection {
    <#
    .SYNOPSIS
        Rename a section in the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SectionIndex
        1-based index of the section to rename.
    .PARAMETER NewName
        New name for the section.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Rename-PowerPointSection -SectionIndex 1 -NewName 'Overview' -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SectionIndex,
        [Parameter(Mandatory)][string]$NewName,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Rename-PowerPointSection'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $oldName = try { $pres.SectionProperties.Name($SectionIndex) } catch { "Section $SectionIndex" }

    if ($PSCmdlet.ShouldProcess($oldName, "Rename section to '$NewName'")) {
        $pres.SectionProperties.Rename($SectionIndex, $NewName)
    }

    $result = [ordered]@{
        status = 'ok'
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Move-PowerPointSection {
    <#
    .SYNOPSIS
        Move a section to a new position in the presentation.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER SectionIndex
        1-based index of the section to move.
    .PARAMETER NewIndex
        1-based target position for the section.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Move-PowerPointSection -SectionIndex 3 -NewIndex 1 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SectionIndex,
        [Parameter(Mandatory)][int]$NewIndex,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Move-PowerPointSection'
    $app  = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    $sectionName = try { $pres.SectionProperties.Name($SectionIndex) } catch { "Section $SectionIndex" }

    if ($PSCmdlet.ShouldProcess($sectionName, "Move section from index $SectionIndex to $NewIndex")) {
        $pres.SectionProperties.Move($SectionIndex, $NewIndex)
    }

    $result = [ordered]@{
        status = 'ok'
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
