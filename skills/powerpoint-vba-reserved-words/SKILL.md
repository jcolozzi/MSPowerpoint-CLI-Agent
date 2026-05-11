---
name: powerpoint-vba-reserved-words
description: >-
  Detect and avoid Microsoft PowerPoint / VBA reserved words in identifiers.
license: CC-BY-4.0
---

# PowerPoint/VBA Reserved Words – Naming Safety Skill

## Purpose
Help developers avoid case-insensitive identifier collisions in PowerPoint VBA.

### Collision Categories

**VBA keywords & operators**
`Dim`, `If`, `Select`, `Function`, `Sub`, `ByVal`, `ByRef`, `And`, `Or`, `Not`, `Set`, `Let`, `New`, `With`, `End`, `Exit`, `GoTo`, `Resume`, `Error`, `On`, `Call`, `Return`, `Nothing`, `Me`, `Optional`, `ParamArray`, etc.

**Built-in VBA function names**
`InStr`, `Left`, `Right`, `Mid`, `Len`, `Date`, `Time`, `Now`, `Format`, `CStr`, `CInt`, `CLng`, `CDbl`, `CBool`, `IsNull`, `IsEmpty`, `IsNumeric`, `Trim`, `UCase`, `LCase`, `Replace`, `Split`, `Join`, `Array`, `UBound`, `LBound`, `Asc`, `Chr`, `RGB`, etc.

**PowerPoint object model names**
`Presentation`, `Presentations`, `Slide`, `Slides`, `SlideRange`, `Shape`, `Shapes`, `ShapeRange`, `TextRange`, `TextFrame`, `TextFrame2`, `Table`, `Chart`, `ChartData`, `Section`, `Design`, `CustomLayout`, `Master`, `SlideMaster`, `SlideLayout`, `SlideShow`, `SlideShowView`, `SlideShowWindow`, `SlideShowSettings`, `SlideShowTransition`, `Effect`, `AnimationSettings`, `AnimationBehavior`, `Sequence`, `TimeLine`, `Placeholder`, `Placeholders`, `HeaderFooter`, `HeadersFooters`, `ColorScheme`, `ColorFormat`, `FillFormat`, `LineFormat`, `ShadowFormat`, `ThreeDFormat`, `PictureFormat`, `TextEffectFormat`, `GroupShapes`, `FreeformBuilder`, `ActionSetting`, `ActionSettings`, `Hyperlink`, `Hyperlinks`, `Comment`, `Comments`, `Tag`, `Tags`, `DocumentProperty`, `CustomDocumentProperties`, `AddIn`, `AddIns`, `Selection`, `View`, `Window`, `Pane`, `Font`, `ParagraphFormat`, `BulletFormat`, `TabStops`, `Column`, `Columns`, `Row`, `Rows`, `Cell`, `CellRange`, `Series`, `SeriesCollection`, `Axis`, `Axes`, `Legend`, `DataLabel`, `DataLabels`, `ChartGroup`, `PrintOptions`, `PageSetup`, `PrintRange`, `PrintRanges`, `Player`, `MediaFormat`, `SmartArt`, `SmartArtLayout`, `SmartArtNode`, `SmartArtNodes`.

**PowerPoint enumeration prefixes**
`pp`, `mso`, `xl` — identifiers starting with these prefixes can collide with PowerPoint/Office constants (e.g., `ppLayoutTitle`, `msoShapeRectangle`, `xlColumnClustered`).

**Special characters/symbols**
Spaces, `'`, `"`, `.`, `!`, `?`, `*`, `+`, `-`, `=`, `<`, `>`, `#`, `%`, `$`, `&`, `@`, `\`, `/`, `^`, `~`, `{}`, `[]`, `()`.

> PowerPoint and VBA treat names **case-insensitively**; reusing these terms as identifiers often leads to compile or runtime errors.

## Procedure

> **Priority:** Complete Steps 1–4 in order. Step 5 is optional and should be considered only when explicitly requested.

### Step 1: Identify Reserved Words (Priority)
1. **Scan identifiers** in the current context (variables, procedure names, module names, shape names, slide names, and embedded VBA).
2. **Flag exact case-insensitive matches** to reserved words and report each occurrence with the specific line number and file name.
3. **Partial-match guidance:** Identifiers that *contain* a reserved word as a complete substring (e.g., `SlideCount`, `ShapeWidth`) should be noted as potential concerns, but only exact or very obvious collisions require immediate action. Use judgment based on context.

### Step 2: Suggest Safe Replacements
4. **Propose descriptive, context-specific names**:
   - `Slide` → `sldCurrent` or `targetSlide`
   - `Shape` → `shpLogo` or `titleShape`
   - `Name` → `ShapeName` / `SlideName`
   - `Text` → `BodyText` / `SlideText`
   - `Table` → `tblData` / `summaryTable`
   - `Chart` → `chtSales` / `revenueChart`
   - `Design` → `slideDesign` / `templateDesign`
   - `Section` → `secIntro` / `presentationSection`
   - `Effect` → `effFadeIn` / `animationEffect`

### Step 3: Apply Naming Conventions (After renaming)
5. **Ensure correct formatting**:
   - Use CamelCase (no spaces or special characters).
   - Apply descriptive prefixes for clarity (see Naming Recommendations below).

### Step 4: Handle Existing Objects
6. **If renaming VBA identifiers** already in use, prefer renaming the identifier rather than using workarounds; clean names prevent subtle bugs and improve readability.

### Step 5: Repository Analysis (Optional)
7. **(Optional) Run a repository scan** when requested to list offenders and propose bulk renames.

## Common Offenders (teach by example)
- PowerPoint objects used as names: `Slide`, `Shape`, `Name`, `Text`, `Table`, `Chart`, `Section`, `Design`, `Layout`, `Master`, `Effect`, `Range`, `Selection`, `View`, `Window`
- Dimension/position collisions: `Index`, `Count`, `Type`, `Value`, `Width`, `Height`, `Left`, `Top`
- VBA functions used as names: `Format`, `Left`, `Right`, `Mid`, `Len`, `Date`, `Time`, `Now`, `Replace`, `Split`, `Join`, `Array`, `RGB`
- Enumeration-like names: `Font`, `Color`, `Fill`, `Line`, `Shadow`, `Cell`, `Row`, `Column`, `Series`, `Axis`

## Naming Recommendations

### Preferred Prefixes (PowerPoint-specific)
| Prefix | Meaning | Example |
|--------|---------|---------|
| `sld` | Slide | `sldTitle`, `sldAgenda` |
| `shp` | Shape | `shpLogo`, `shpArrow` |
| `tf` | TextFrame | `tfBody`, `tfSubtitle` |
| `tr` | TextRange | `trHeading`, `trBullets` |
| `cht` | Chart | `chtRevenue`, `chtPipeline` |
| `tbl` | Table | `tblSchedule`, `tblMetrics` |
| `ani` | Animation | `aniEntrance`, `aniFadeIn` |
| `eff` | Effect | `effWipe`, `effZoom` |
| `sec` | Section | `secIntro`, `secAppendix` |
| `lyt` | Layout | `lytTwoColumn`, `lytBlank` |
| `mst` | Master | `mstCorporate`, `mstDefault` |
| `img` | Image | `imgHero`, `imgBackground` |
| `lnk` | Hyperlink | `lnkWebsite`, `lnkEmail` |
| `prs` | Presentation | `prsSource`, `prsTarget` |

### General Rules
- ✅ Prefer descriptive, specific names: `sldExecutiveSummary`, `shpCompanyLogo`, `chtQuarterlySales`.
- ✅ Use CamelCase; consider type/object prefixes for clarity.
- ✅ Use `Option Explicit` in every VBA module.
- ❌ Avoid reserved words and special characters in any identifier.
- ❌ Avoid generic names like `Data`, `Info`, `Temp` when they collide with engine terms.
- ❌ Avoid names starting with `pp`, `mso`, or `xl` — these collide with PowerPoint/Office enumeration constants.

## Example Prompts (to trigger this skill)
- "Scan this VBA module and flag any **PowerPoint/VBA reserved-word** variable names."
- "Suggest safe replacements for variables named `Slide` and `Shape` in my macro."
- "Check my PowerPoint VBA code for identifiers that collide with the PowerPoint object model."
- "Rename my shape variables to follow PowerPoint naming conventions."

## References
- Microsoft: [PowerPoint VBA object model reference](https://learn.microsoft.com/en-us/office/vba/api/overview/powerpoint)
- Microsoft: [VBA language reference](https://learn.microsoft.com/en-us/office/vba/api/overview/language-reference)
- Microsoft: [PowerPoint enumerations](https://learn.microsoft.com/en-us/office/vba/api/powerpoint(enumerations))
- VBA language keywords & reserved identifiers
