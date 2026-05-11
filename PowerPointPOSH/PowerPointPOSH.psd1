@{
    RootModule        = 'PowerPointPOSH.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890'
    Author            = 'PowerPointPOSH'
    Description       = 'PowerShell PowerPoint Presentation Automation via COM — 127 functions for presentation, slide, shape, text, table, chart, image/media, animation, transition, slideshow, formatting, export, master/layout, metadata, print, VBE, SmartArt, hyperlink, and section operations'
    PowerShellVersion = '5.1'

    FunctionsToExport = @(
        # Application (3)
        'Get-PowerPointApplicationInfo'
        'Set-PowerPointOption'
        'Get-PowerPointTip'

        # Presentation (8)
        'Open-PowerPointPresentation'
        'New-PowerPointPresentation'
        'Save-PowerPointPresentation'
        'Close-PowerPointPresentation'
        'Get-PowerPointPresentationInfo'
        'Copy-PowerPointPresentation'
        'Convert-PowerPointPresentation'
        'Repair-PowerPointPresentation'

        # Slide (10)
        'Get-PowerPointSlide'
        'New-PowerPointSlide'
        'Remove-PowerPointSlide'
        'Copy-PowerPointSlide'
        'Move-PowerPointSlide'
        'Set-PowerPointSlideLayout'
        'Get-PowerPointSlideNotes'
        'Set-PowerPointSlideNotes'
        'Set-PowerPointSlideBackground'
        'Get-PowerPointSlidePlaceholders'

        # Shape (14)
        'Get-PowerPointShape'
        'Add-PowerPointShape'
        'Add-PowerPointTextBox'
        'Add-PowerPointLine'
        'Remove-PowerPointShape'
        'Set-PowerPointShapeProperties'
        'Copy-PowerPointShape'
        'Group-PowerPointShapes'
        'Ungroup-PowerPointShapes'
        'Set-PowerPointShapeZOrder'
        'Get-PowerPointShapeList'
        'Find-PowerPointShape'
        'Align-PowerPointShapes'
        'Distribute-PowerPointShapes'

        # Text (8)
        'Get-PowerPointText'
        'Set-PowerPointText'
        'Format-PowerPointTextFont'
        'Format-PowerPointTextParagraph'
        'Add-PowerPointBullet'
        'Find-PowerPointText'
        'Set-PowerPointTextReplace'
        'Get-PowerPointTextAll'

        # Table (7)
        'Add-PowerPointTable'
        'Get-PowerPointTable'
        'Set-PowerPointTableCell'
        'Get-PowerPointTableCell'
        'Format-PowerPointTableCell'
        'Add-PowerPointTableRow'
        'Remove-PowerPointTableRow'

        # Chart (6)
        'Add-PowerPointChart'
        'Get-PowerPointChart'
        'Set-PowerPointChart'
        'Set-PowerPointChartData'
        'Set-PowerPointChartSeries'
        'Export-PowerPointChart'

        # ImageMedia (7)
        'Add-PowerPointImage'
        'Add-PowerPointAudio'
        'Add-PowerPointVideo'
        'Get-PowerPointMedia'
        'Set-PowerPointMediaProperties'
        'Remove-PowerPointMedia'
        'Export-PowerPointSlideImage'

        # Animation (7)
        'Add-PowerPointAnimation'
        'Get-PowerPointAnimation'
        'Remove-PowerPointAnimation'
        'Set-PowerPointAnimationTiming'
        'Set-PowerPointAnimationOrder'
        'Clear-PowerPointAnimations'
        'Copy-PowerPointAnimation'

        # Transition (4)
        'Set-PowerPointTransition'
        'Get-PowerPointTransition'
        'Remove-PowerPointTransition'
        'Copy-PowerPointTransition'

        # SlideShow (6)
        'Start-PowerPointSlideShow'
        'Stop-PowerPointSlideShow'
        'Set-PowerPointSlideShowSettings'
        'Get-PowerPointSlideShowInfo'
        'Step-PowerPointSlideShow'
        'Set-PowerPointPresenterView'

        # Formatting (7)
        'Set-PowerPointShapeFill'
        'Set-PowerPointShapeLine'
        'Set-PowerPointShapeShadow'
        'Set-PowerPointShapeEffect'
        'Set-PowerPointThemeColor'
        'Set-PowerPointShapeSize'
        'Set-PowerPointShapePosition'

        # Export (6)
        'Export-PowerPointToPdf'
        'Export-PowerPointToImages'
        'Export-PowerPointToVideo'
        'Export-PowerPointToHtml'
        'Export-PowerPointSlide'
        'Convert-PowerPointFormat'

        # MasterLayout (5)
        'Get-PowerPointSlideMaster'
        'Get-PowerPointSlideLayout'
        'Set-PowerPointSlideMaster'
        'New-PowerPointCustomLayout'
        'Get-PowerPointPlaceholder'

        # Metadata (7)
        'Get-PowerPointDocumentProperty'
        'Set-PowerPointDocumentProperty'
        'Get-PowerPointComment'
        'Add-PowerPointComment'
        'Remove-PowerPointComment'
        'Get-PowerPointTag'
        'Set-PowerPointTag'

        # Print (3)
        'Set-PowerPointPageSetup'
        'Get-PowerPointPageSetup'
        'Invoke-PowerPointPrint'

        # VBE (7)
        'Get-PowerPointVbaCode'
        'Set-PowerPointVbaCode'
        'Add-PowerPointVbaModule'
        'Remove-PowerPointVbaModule'
        'Get-PowerPointVbaModuleList'
        'Find-PowerPointVbaCode'
        'Invoke-PowerPointVbaMacro'

        # SmartArt (4)
        'Add-PowerPointSmartArt'
        'Get-PowerPointSmartArt'
        'Set-PowerPointSmartArtLayout'
        'Set-PowerPointSmartArtNode'

        # Hyperlink (3)
        'Add-PowerPointHyperlink'
        'Get-PowerPointHyperlink'
        'Remove-PowerPointHyperlink'

        # Section (5)
        'Get-PowerPointSection'
        'New-PowerPointSection'
        'Remove-PowerPointSection'
        'Rename-PowerPointSection'
        'Move-PowerPointSection'
    )

    CmdletsToExport   = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
}
