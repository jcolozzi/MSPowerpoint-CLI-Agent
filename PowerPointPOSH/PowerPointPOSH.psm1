<#
.SYNOPSIS
    PowerPointPOSH — PowerShell PowerPoint Presentation Automation

.DESCRIPTION
    127 public functions for presentation, slide, shape, text, table, chart,
    image/media, animation, transition, slideshow, formatting, export,
    master/layout, metadata, print, VBE, SmartArt, hyperlink, and section
    operations.
    No MCP server needed — AI agents call functions directly via terminal.

    Usage:
        Import-Module .\PowerPointPOSH\PowerPointPOSH.psd1 -Force
        Get-PowerPointSlide -PresentationPath "C:\deck.pptx" -SlideIndex 1 -AsJson
        Close-PowerPointPresentation                  # release COM

.NOTES
    Requires: Windows + Microsoft PowerPoint (Microsoft 365)
    PowerShell: 5.1+ or PowerShell 7+
#>

# ═══════════════════════════════════════════════════════════════════════════
# CONSTANTS & TYPE MAPS
# ═══════════════════════════════════════════════════════════════════════════

# PpSaveAsFileType
$script:PPT_FILE_FORMAT = @{
    pptx = 24; pptm = 25; potx = 26; potm = 27; ppsx = 28; ppsm = 29
    ppt = 1; pps = 7; pot = 5
    pdf = 32; xps = 33
    png = 18; jpg = 17; gif = 16; bmp = 19; tif = 21; emf = 23
    mp4 = 39; wmv = 37
    rtf = 6; odp = 35
    animatedgif = 40
}

# PpSlideLayout
$script:PPT_SLIDE_LAYOUT = @{
    title = 1; text = 2; twoColumnText = 3; table = 4
    textAndChart = 5; chartAndText = 6; orgchart = 7; chart = 8
    textAndClipArt = 9; clipArtAndText = 10; titleOnly = 11; blank = 12
    textAndObject = 13; objectAndText = 14; largeObject = 15; object = 16
    textAndMediaClip = 17; mediaClipAndText = 18; objectOverText = 19
    textOverObject = 20; twoObjects = 29; twoObjectsAndText = 22
    twoObjectsOverText = 23; fourObjects = 24; verticalText = 25
    clipArtAndVerticalText = 26; verticalTitleAndText = 27
    verticalTitleAndTextOverChart = 28; twoObjectsAndObject = 30
    objectAndTwoObjects = 31; custom = 32; sectionHeader = 33
    comparison = 34; contentWithCaption = 35; pictureWithCaption = 36
}

# MsoAutoShapeType (common shapes)
$script:PPT_SHAPE_TYPE = @{
    rectangle = 1; parallelogram = 2; trapezoid = 3; diamond = 4
    roundedRectangle = 5; octagon = 6; triangle = 7; rightTriangle = 8
    oval = 9; hexagon = 10; cross = 11; star5 = 12
    rightArrow = 33; leftArrow = 34; upArrow = 35; downArrow = 36
    leftRightArrow = 37; upDownArrow = 38; chevron = 52
    heart = 21; lightningBolt = 22; sun = 23; moon = 24
    smileyFace = 17; donut = 18; noSymbol = 19; cube = 14
    can = 13; foldedCorner = 16; bevel = 15
    callout1 = 41; callout2 = 42; callout3 = 43; callout4 = 44
    cloud = 179; frame = 158; plaque = 28; star4 = 187; star6 = 188
    star8 = 58; star10 = 189; star12 = 190; star16 = 94; star24 = 95; star32 = 96
    roundedRectangleCallout = 105; ovalCallout = 107
    cloudCallout = 108
}

# MsoAnimEffect (common animation effects)
$script:PPT_ANIM_EFFECT = @{
    appear = 1; blinds = 3; box = 4; checkerboard = 5
    circle = 6; crawl = 7; diamond = 8; dissolve = 9
    fadeIn = 10; flash = 11; flyIn = 2; peek = 12
    plus = 13; randomBars = 14; spiral = 15; split = 16
    stretch = 17; strips = 18; swivel = 19; wedge = 20
    wheel = 21; wipe = 22; zoom = 23; randomEffects = 24
    boomerang = 25; bounce = 26; colorPulse = 27; colorWave = 28
    colorBlend = 29; complementaryColor = 30; complementaryColor2 = 31
    grow = 43; shrink = 44; spin = 49; transparency = 50
    fadeOut = 63; float = 64; pinwheel = 72; wave = 82
}

# PpEntryEffect (slide transitions - common subset)
$script:PPT_TRANSITION = @{
    none = 0; blindsHorizontal = 769; blindsVertical = 770
    boxIn = 3074; boxOut = 3073; checkerboardAcross = 1025
    checkerboardDown = 1026; combHorizontal = 3847; combVertical = 3848
    coverDown = 1284; coverLeft = 1281; coverRight = 1283; coverUp = 1282
    cut = 257; cutThroughBlack = 258; dissolve = 1537
    fade = 1793; fadeSmoothly = 3849; flashOnce = 3850
    newsflash = 3851; plus = 3852; pushDown = 3853; pushLeft = 3854
    pushRight = 3855; pushUp = 3856; random = 513
    splitHorizontalIn = 3137; splitHorizontalOut = 3138
    splitVerticalIn = 3139; splitVerticalOut = 3140
    stripsDownLeft = 2050; stripsDownRight = 2051
    stripsLeftDown = 2052; stripsLeftUp = 2053
    stripsRightDown = 2054; stripsRightUp = 2055
    stripsUpLeft = 2056; stripsUpRight = 2057
    uncoverDown = 2564; uncoverLeft = 2561; uncoverRight = 2563; uncoverUp = 2562
    wipeDown = 1796; wipeLeft = 1793; wipeRight = 1795; wipeUp = 1794
    appear = 3844; circleOut = 3845
}

# PpPlaceholderType
$script:PPT_PLACEHOLDER = @{
    title = 1; body = 2; centerTitle = 3; subtitle = 4
    dateTime = 6; slideNumber = 7; footer = 8; header = 9
    object = 10; chart = 11; table = 12; clipArt = 13
    orgChart = 14; media = 16; verticalBody = 17; verticalObject = 18
    verticalTitle = 19; bitmap = 20; picture = 21
}

# PpParagraphAlignment
$script:PPT_ALIGN = @{
    left = 1; center = 2; right = 3; justify = 4; distribute = 5
    justifyLow = 6; thaiDistribute = 7
}

# MsoFillType
$script:PPT_FILL_TYPE = @{
    solid = 1; patterned = 2; gradient = 3; textured = 4
    background = 5; picture = 6
}

# MsoLineDashStyle
$script:PPT_LINE_DASH = @{
    solid = 1; squareDot = 2; roundDot = 3; dash = 4
    dashDot = 5; dashDotDot = 6; longDash = 7; longDashDot = 8
    longDashDotDot = 9; sysDash = 11; sysDot = 12; sysDashDot = 13
}

# PpAlertLevel
$script:PPT_ALERT_LEVEL = @{
    none = 0   # ppAlertsNone
    all  = 1   # ppAlertsAll
}

# PpMediaType
$script:PPT_MEDIA_TYPE = @{
    other = 1; sound = 2; movie = 3; mixed = -2
}

# PpSlideSizeType
$script:PPT_SLIDE_SIZE = @{
    onScreen = 1; letterPaper = 2; a4Paper = 5; slide35mm = 6
    overhead = 7; banner = 8; custom = 9; ledgerPaper = 10
    a3Paper = 11; b4ISOPaper = 12; b5ISOPaper = 13; b4JISPaper = 14
    b5JISPaper = 15; hagakiCard = 16; onScreen16x9 = 17; onScreen16x10 = 18
    widescreen = 32
}

# PpWindowState
$script:PPT_WINDOW_STATE = @{
    normal = 1; minimized = 2; maximized = 3
}

# PpPrintOutputType
$script:PPT_PRINT_OUTPUT = @{
    slides = 1; twoSlideHandouts = 2; threeSlideHandouts = 3
    sixSlideHandouts = 4; notesPages = 5; outline = 6
    fourSlideHandouts = 7; nineSlideHandouts = 8
}

# MsoZOrderCmd
$script:PPT_ZORDER = @{
    bringToFront = 0; sendToBack = 1; bringForward = 2; sendBackward = 3
    bringInFrontOfText = 4; sendBehindText = 5
}

# MsoAlignCmd — shape alignment
$script:PPT_ALIGN_CMD = @{
    alignLefts = 0; alignCenters = 1; alignRights = 2
    alignTops = 3; alignMiddles = 4; alignBottoms = 5
}

# MsoDistributeCmd
$script:PPT_DISTRIBUTE = @{
    distributeHorizontally = 0; distributeVertically = 1
}

# PpFixedFormatType
$script:PPT_FIXED_FORMAT = @{
    pdf = 2   # ppFixedFormatTypePDF
    xps = 1   # ppFixedFormatTypeXPS
}

# MsoTriState (for boolean-like COM props)
$script:PPT_TRISTATE = @{
    true   = -1  # msoTrue
    false  = 0   # msoFalse
    mixed  = -2  # msoTriStateMixed
    toggle = -3  # msoTriStateToggle
}

# PpSelectionType
$script:PPT_SELECTION_TYPE = @{
    none = 0; slides = 1; shapes = 2; text = 3
}

# PpActionType
$script:PPT_ACTION_TYPE = @{
    none = 0; nextSlide = 1; previousSlide = 2; firstSlide = 3
    lastSlide = 4; lastSlideViewed = 5; endShow = 6; hyperlink = 7
    runMacro = 8; runProgram = 9; namedSlideShow = 10; oleVerb = 11
    play = 12
}

# MsoAnimType
$script:PPT_ANIM_TYPE = @{
    none = 0; motion = 1; color = 2; scale = 3; rotation = 4
    property = 5; command = 6; filter = 7; set = 8; mixed = -2
}

# MsoTextOrientation
$script:PPT_TEXT_ORIENTATION = @{
    horizontal = 1; verticalFarEast = 6; vertical = 5
    upward = 2; downward = 3; mixed = -2
}

# ═══════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════

$script:PowerPointSession = @{
    App              = $null   # COM PowerPoint.Application object
    PresentationPath = $null   # Currently open presentation path (resolved)
    OwnsApp          = $false  # $true when we created the COM instance; controls Quit() on exit
}

# ═══════════════════════════════════════════════════════════════════════════
# DOT-SOURCE ALL SUB-FILES
# ═══════════════════════════════════════════════════════════════════════════

# Private helpers first (session, utilities)
foreach ($file in (Get-ChildItem -Path "$PSScriptRoot\Private\*.ps1" -ErrorAction SilentlyContinue)) {
    . $file.FullName
}

# Public domain files
foreach ($file in (Get-ChildItem -Path "$PSScriptRoot\Public\*.ps1" -ErrorAction SilentlyContinue)) {
    . $file.FullName
}

# ═══════════════════════════════════════════════════════════════════════════
# CLEANUP ON EXIT
# ═══════════════════════════════════════════════════════════════════════════

Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    if ($null -ne $script:PowerPointSession -and $null -ne $script:PowerPointSession.App) {
        # Only Quit if we created the instance
        if ($script:PowerPointSession.OwnsApp) {
            try { $script:PowerPointSession.App.DisplayAlerts = 0 } catch {}  # ppAlertsNone
            try { $script:PowerPointSession.App.Quit() } catch {}
        }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:PowerPointSession.App) } catch {}
        $script:PowerPointSession.App              = $null
        $script:PowerPointSession.PresentationPath = $null
        $script:PowerPointSession.OwnsApp          = $false
    }
} | Out-Null

# ═══════════════════════════════════════════════════════════════════════════
# LOADED
# ═══════════════════════════════════════════════════════════════════════════
Write-Host 'PowerPointPOSH module loaded. Use Close-PowerPointPresentation to release COM when done.' -ForegroundColor Cyan
