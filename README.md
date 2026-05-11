# PowerPointPOSH — PowerShell PowerPoint Presentation Automation

COM automation of Microsoft PowerPoint presentations.  
**127 public functions** across 20 categories — no MCP server needed.  
AI agents call functions directly via the terminal.

---

## Quick Start

```powershell
Import-Module .\PowerPointPOSH\PowerPointPOSH.psd1 -Force
Open-PowerPointPresentation -PresentationPath "C:\deck.pptx" -AsJson
Get-PowerPointSlide -AsJson
New-PowerPointSlide -Layout 'blank' -AsJson
Add-PowerPointShape -SlideIndex 2 -ShapeType 'rectangle' -Left 100 -Top 100 -Width 200 -Height 100 -AsJson
Set-PowerPointText -SlideIndex 2 -ShapeName "Rectangle 1" -Text "Hello World" -AsJson
Export-PowerPointToPdf -FilePath "C:\output.pdf" -AsJson
Close-PowerPointPresentation
```

---

## Requirements

| Requirement | Version |
|---|---|
| OS | Windows |
| PowerPoint | Microsoft 365 (desktop) |
| PowerShell | 5.1+ |

---

## Function Categories

| # | Category | Count | Description |
|---|---|---|---|
| 1 | Application | 3 | App info, settings, tips |
| 2 | Presentation | 8 | Open, new, save, close, info, copy, convert, repair |
| 3 | Slide | 10 | CRUD, layout, notes, background, placeholders |
| 4 | Shape | 14 | Add/remove/modify shapes, group, align, z-order |
| 5 | Text | 8 | Read/write/format text, find/replace, bullets |
| 6 | Table | 7 | Add tables, read/write cells, format |
| 7 | Chart | 6 | Add/modify charts, set data |
| 8 | Image & Media | 7 | Pictures, audio, video, export slide images |
| 9 | Animation | 7 | Add/remove/configure animations |
| 10 | Transition | 4 | Set/get/remove slide transitions |
| 11 | SlideShow | 6 | Start/stop/navigate presentations |
| 12 | Formatting | 7 | Fill, line, shadow, effects, size, position |
| 13 | Export | 6 | PDF, images, video, HTML |
| 14 | Master & Layout | 5 | Slide masters, custom layouts |
| 15 | Metadata | 7 | Properties, comments, tags |
| 16 | Print | 3 | Page setup, print |
| 17 | VBE | 7 | VBA code read/write/execute (macro-enabled only) |
| 18 | SmartArt | 4 | Add/modify SmartArt graphics |
| 19 | Hyperlink | 3 | Add/get/remove hyperlinks |
| 20 | Section | 5 | Manage presentation sections |

**Total: 127 functions**

---

## Full Function Reference

### Application (3)

| Function | Description |
|---|---|
| `Get-PowerPointApplication` | Returns running PowerPoint app info (version, build, visible state) |
| `Set-PowerPointApplication` | Configures app-level settings (visible, display alerts) |
| `Get-PowerPointTips` | Lists tips and gotchas for PowerPoint COM automation |

### Presentation (8)

| Function | Description |
|---|---|
| `Open-PowerPointPresentation` | Opens an existing presentation file |
| `New-PowerPointPresentation` | Creates a new blank presentation |
| `Save-PowerPointPresentation` | Saves the active presentation |
| `Close-PowerPointPresentation` | Closes the active presentation and optionally quits the app |
| `Get-PowerPointPresentationInfo` | Returns presentation metadata (slide count, size, path) |
| `Copy-PowerPointPresentation` | Saves a copy of the active presentation to a new path |
| `Convert-PowerPointPresentation` | Converts between formats (pptx, pptm, ppt, potx, ppsx) |
| `Repair-PowerPointPresentation` | Attempts to repair a corrupted presentation |

### Slide (10)

| Function | Description |
|---|---|
| `Get-PowerPointSlide` | Lists all slides with index, layout, name, and shape count |
| `New-PowerPointSlide` | Adds a new slide with specified layout |
| `Remove-PowerPointSlide` | Deletes a slide by index |
| `Copy-PowerPointSlide` | Duplicates a slide within or across presentations |
| `Move-PowerPointSlide` | Moves a slide to a new position |
| `Set-PowerPointSlideLayout` | Changes the layout of an existing slide |
| `Get-PowerPointSlideNotes` | Reads the notes pane text for a slide |
| `Set-PowerPointSlideNotes` | Sets or appends text in the notes pane |
| `Set-PowerPointSlideBackground` | Sets slide background (solid, gradient, image) |
| `Get-PowerPointPlaceholder` | Lists placeholders on a slide with type and position |

### Shape (14)

| Function | Description |
|---|---|
| `Add-PowerPointShape` | Adds an AutoShape (rectangle, oval, arrow, etc.) |
| `Get-PowerPointShape` | Lists shapes on a slide with name, type, position, size |
| `Remove-PowerPointShape` | Deletes a shape by name or index |
| `Set-PowerPointShapePosition` | Moves a shape to specified left/top coordinates |
| `Set-PowerPointShapeSize` | Resizes a shape (width, height) |
| `Copy-PowerPointShape` | Duplicates a shape on the same or different slide |
| `Group-PowerPointShape` | Groups multiple shapes into a single group |
| `Ungroup-PowerPointShape` | Ungroups a grouped shape |
| `Set-PowerPointShapeZOrder` | Changes z-order (bring to front, send to back, etc.) |
| `Align-PowerPointShape` | Aligns shapes relative to each other or the slide |
| `Distribute-PowerPointShape` | Distributes shapes evenly (horizontal or vertical) |
| `Rotate-PowerPointShape` | Rotates a shape by angle or flip direction |
| `Lock-PowerPointShape` | Locks a shape to prevent accidental editing |
| `Rename-PowerPointShape` | Renames a shape |

### Text (8)

| Function | Description |
|---|---|
| `Get-PowerPointText` | Reads text content from a shape or placeholder |
| `Set-PowerPointText` | Writes text to a shape or placeholder |
| `Format-PowerPointText` | Applies font formatting (bold, italic, size, color, font name) |
| `Add-PowerPointTextBox` | Adds a text box shape with specified text |
| `Find-PowerPointText` | Searches for text across slides |
| `Replace-PowerPointText` | Find and replace text across slides |
| `Set-PowerPointBullet` | Configures bullet/numbering for a text frame |
| `Set-PowerPointParagraphFormat` | Sets paragraph alignment, spacing, and indentation |

### Table (7)

| Function | Description |
|---|---|
| `Add-PowerPointTable` | Inserts a table with specified rows and columns |
| `Get-PowerPointTableCell` | Reads cell values from a table |
| `Set-PowerPointTableCell` | Writes a value to a table cell |
| `Get-PowerPointTable` | Returns full table data as a structured object |
| `Format-PowerPointTableCell` | Formats cell text (font, size, color, alignment) |
| `Set-PowerPointTableStyle` | Applies a built-in table style |
| `Add-PowerPointTableRow` | Adds a row to an existing table |

### Chart (6)

| Function | Description |
|---|---|
| `Add-PowerPointChart` | Inserts a chart (column, bar, line, pie, etc.) |
| `Get-PowerPointChart` | Reads chart configuration and data range |
| `Set-PowerPointChartData` | Updates chart data from arrays or CSV |
| `Set-PowerPointChartStyle` | Applies chart style, title, legend, axis formatting |
| `Set-PowerPointChartType` | Changes chart type of an existing chart |
| `Export-PowerPointChart` | Exports a chart as an image file |

### Image & Media (7)

| Function | Description |
|---|---|
| `Add-PowerPointImage` | Inserts an image from file path |
| `Add-PowerPointAudio` | Embeds an audio clip |
| `Add-PowerPointVideo` | Embeds a video clip |
| `Set-PowerPointMediaPlayback` | Configures media playback settings (autoplay, loop) |
| `Get-PowerPointMedia` | Lists media objects on a slide |
| `Export-PowerPointSlideImage` | Exports a slide as a PNG/JPG image |
| `Set-PowerPointImageCrop` | Crops an image shape |

### Animation (7)

| Function | Description |
|---|---|
| `Add-PowerPointAnimation` | Adds an animation effect to a shape |
| `Get-PowerPointAnimation` | Lists animations on a slide with sequence and timing |
| `Remove-PowerPointAnimation` | Removes an animation from a shape |
| `Set-PowerPointAnimationTiming` | Configures animation duration, delay, and repeat |
| `Set-PowerPointAnimationTrigger` | Sets trigger type (on click, with previous, after previous) |
| `Set-PowerPointAnimationOrder` | Reorders animation sequence |
| `Copy-PowerPointAnimation` | Copies animation settings between shapes |

### Transition (4)

| Function | Description |
|---|---|
| `Set-PowerPointTransition` | Applies a transition effect to a slide |
| `Get-PowerPointTransition` | Reads current transition settings for a slide |
| `Remove-PowerPointTransition` | Removes transition from a slide |
| `Set-PowerPointTransitionTiming` | Configures transition duration and advance settings |

### SlideShow (6)

| Function | Description |
|---|---|
| `Start-PowerPointSlideShow` | Starts the slide show from a specified slide |
| `Stop-PowerPointSlideShow` | Ends the running slide show |
| `Set-PowerPointSlideShowSettings` | Configures show type, range, and looping |
| `Invoke-PowerPointSlideShowNavigate` | Navigates to a specific slide during the show |
| `Get-PowerPointSlideShowStatus` | Returns current slide show state and position |
| `Set-PowerPointSlideShowTimings` | Sets rehearsed timings for auto-advance |

### Formatting (7)

| Function | Description |
|---|---|
| `Set-PowerPointFill` | Sets shape fill (solid, gradient, pattern, picture) |
| `Set-PowerPointLine` | Configures shape outline (color, weight, dash style) |
| `Set-PowerPointShadow` | Applies shadow effect to a shape |
| `Set-PowerPointEffect` | Applies 3D, reflection, or glow effects |
| `Set-PowerPointShapeStyle` | Applies a built-in shape style |
| `Get-PowerPointFormatting` | Reads current formatting properties of a shape |
| `Reset-PowerPointFormatting` | Resets shape formatting to layout defaults |

### Export (6)

| Function | Description |
|---|---|
| `Export-PowerPointToPdf` | Exports presentation as PDF |
| `Export-PowerPointToImages` | Exports all slides as image files |
| `Export-PowerPointToVideo` | Exports presentation as MP4 video |
| `Export-PowerPointToHtml` | Exports presentation as HTML |
| `Export-PowerPointSlide` | Exports a single slide in a specified format |
| `Export-PowerPointHandout` | Exports a handout layout (multiple slides per page) |

### Master & Layout (5)

| Function | Description |
|---|---|
| `Get-PowerPointSlideMaster` | Lists slide masters in the presentation |
| `Get-PowerPointSlideLayout` | Lists available layouts from a slide master |
| `New-PowerPointSlideLayout` | Creates a custom layout |
| `Set-PowerPointSlideMasterElement` | Modifies elements on a slide master |
| `Copy-PowerPointSlideMaster` | Copies a slide master from another presentation |

### Metadata (7)

| Function | Description |
|---|---|
| `Get-PowerPointProperty` | Reads built-in document properties (title, author, etc.) |
| `Set-PowerPointProperty` | Sets built-in or custom document properties |
| `Get-PowerPointComment` | Lists comments on a slide |
| `Add-PowerPointComment` | Adds a comment to a slide |
| `Remove-PowerPointComment` | Removes a comment |
| `Get-PowerPointTag` | Reads custom tags on slides or shapes |
| `Set-PowerPointTag` | Sets custom tags on slides or shapes |

### Print (3)

| Function | Description |
|---|---|
| `Get-PowerPointPageSetup` | Returns slide size and orientation |
| `Set-PowerPointPageSetup` | Sets slide size (standard, widescreen, custom) |
| `Invoke-PowerPointPrint` | Prints the presentation |

### VBE (7)

| Function | Description |
|---|---|
| `Get-PowerPointVbaCode` | Reads VBA module code |
| `Set-PowerPointVbaCode` | Writes or replaces VBA module code |
| `Add-PowerPointVbaModule` | Adds a new VBA module |
| `Remove-PowerPointVbaModule` | Removes a VBA module |
| `Invoke-PowerPointVbaMacro` | Executes a VBA macro by name |
| `Get-PowerPointVbaReference` | Lists VBA project references |
| `Add-PowerPointVbaReference` | Adds a VBA reference by GUID or file path |

### SmartArt (4)

| Function | Description |
|---|---|
| `Add-PowerPointSmartArt` | Inserts a SmartArt graphic |
| `Get-PowerPointSmartArt` | Reads SmartArt layout and node data |
| `Set-PowerPointSmartArtNode` | Modifies text or properties of a SmartArt node |
| `Set-PowerPointSmartArtLayout` | Changes the SmartArt layout type |

### Hyperlink (3)

| Function | Description |
|---|---|
| `Add-PowerPointHyperlink` | Adds a hyperlink to a shape or text range |
| `Get-PowerPointHyperlink` | Lists hyperlinks on a slide |
| `Remove-PowerPointHyperlink` | Removes a hyperlink from a shape |

### Section (5)

| Function | Description |
|---|---|
| `Get-PowerPointSection` | Lists all sections in the presentation |
| `Add-PowerPointSection` | Adds a new section at a slide index |
| `Remove-PowerPointSection` | Removes a section (slides remain) |
| `Rename-PowerPointSection` | Renames a section |
| `Move-PowerPointSection` | Moves a section to a new position |

---

## Architecture

```
PowerPointPOSH/
├── PowerPointPOSH.psd1          # Module manifest — exports 127 functions
├── PowerPointPOSH.psm1          # Root module — constants, session state, dot-sourcing, cleanup
├── Private/
│   ├── Session.ps1              # COM session management (launch, attach, release)
│   └── Utilities.ps1            # Value conversion, enum mapping, JSON output formatting
└── Public/
    ├── ApplicationOps.ps1       # 3 functions
    ├── PresentationOps.ps1      # 8 functions
    ├── SlideOps.ps1             # 10 functions
    ├── ShapeOps.ps1             # 14 functions
    ├── TextOps.ps1              # 8 functions
    ├── TableOps.ps1             # 7 functions
    ├── ChartOps.ps1             # 6 functions
    ├── ImageMediaOps.ps1        # 7 functions
    ├── AnimationOps.ps1         # 7 functions
    ├── TransitionOps.ps1        # 4 functions
    ├── SlideShowOps.ps1         # 6 functions
    ├── FormattingOps.ps1        # 7 functions
    ├── ExportOps.ps1            # 6 functions
    ├── MasterLayoutOps.ps1      # 5 functions
    ├── MetadataOps.ps1          # 7 functions
    ├── PrintOps.ps1             # 3 functions
    ├── VbeOps.ps1               # 7 functions
    ├── SmartArtOps.ps1          # 4 functions
    ├── HyperlinkOps.ps1         # 3 functions
    └── SectionOps.ps1           # 5 functions
```

### Session Management

The module maintains a single `$script:PowerPointSession` hashtable:

- **App** — The PowerPoint `Application` COM object
- **Presentation** — The active `Presentation` COM object
- **PresentationPath** — Full path to the open file

`Open-PowerPointPresentation` creates the session. `Close-PowerPointPresentation` releases COM references and optionally quits the app. The module's `OnRemove` handler calls `[System.Runtime.InteropServices.Marshal]::ReleaseComObject()` for deterministic cleanup.

### COM Cleanup

All public functions wrap COM access in `try/finally` blocks. Intermediate COM objects (slides, shapes, ranges) are released after use. `Remove-Module PowerPointPOSH` triggers the cleanup handler automatically.

---

## Testing

```powershell
Invoke-Pester .\Tests\ -Output Detailed
```

The test suite includes:

- **PowerPointPOSH.Module.Tests.ps1** — Module-level tests (manifest, exports, function naming)
- **20 domain test files** — One per category (e.g., `SlideOps.Tests.ps1`, `ShapeOps.Tests.ps1`)

Tests mock the COM layer and validate parameter binding, error handling, and JSON output structure.

---

## Agent Integration

AI agents interact with PowerPointPOSH through terminal commands:

```powershell
# Agent imports the module
Import-Module .\PowerPointPOSH\PowerPointPOSH.psd1 -Force

# All commands return structured JSON with -AsJson
$slides = Get-PowerPointSlide -AsJson | ConvertFrom-Json

# Agents parse JSON to make decisions
$emptySlides = $slides | Where-Object { $_.ShapeCount -eq 0 }
```

Key patterns for agents:
- **Always use `-AsJson`** — Returns structured JSON instead of PowerShell objects
- **Absolute paths** — Use full paths for file operations
- **One presentation at a time** — The session model supports a single active presentation
- **Error responses** — Failures return JSON with `success: false` and an error message

---

## Known Limitations

- **Single-presentation model** — Only one presentation can be open per session
- **Desktop PowerPoint required** — COM automation requires the full desktop application; PowerPoint Online is not supported
- **VBE Trust Center** — VBA operations require "Trust access to the VBA project object model" to be enabled in Trust Center settings
- **No concurrent access** — COM is single-threaded; do not run multiple PowerPointPOSH sessions simultaneously
- **Large media files** — Embedding large videos may be slow via COM; consider linking instead
