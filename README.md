# MSPowerPoint-CLI-Agent

> **Automate Microsoft PowerPoint from plain English inside VS Code — no MCP server, no Python, no extra processes.**

![Platform: Windows](https://img.shields.io/badge/platform-Windows-blue?logo=windows)
![PowerShell: 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![VS Code](https://img.shields.io/badge/VS%20Code-GitHub%20Copilot%20Chat-blueviolet?logo=visual-studio-code)
![Functions: 127](https://img.shields.io/badge/functions-127-brightgreen)
![Module Version](https://img.shields.io/badge/version-1.0.0-orange)
![License: MIT](https://img.shields.io/badge/license-MIT-green)

## What is this?

**MSPowerPoint-CLI-Agent** is a VS Code agent (powered by GitHub Copilot Chat) that lets you talk to Microsoft PowerPoint in plain language. You describe what you want and the agent translates it into PowerShell commands that manipulate your `.pptx` / `.pptm` presentation live via COM — no manual VBA editing required.

```text
You:   "Add a new blank slide and put a rectangle with 'Hello World' on it"
Agent: → New-PowerPointSlide → Add-PowerPointShape → Set-PowerPointText → confirms success
```

The **PowerPointPOSH** module (included) provides **127 public functions** across 20 categories covering presentations, slides, shapes, text, tables, charts, media, animations, transitions, slide shows, formatting, export, masters, metadata, VBA/VBE, SmartArt, hyperlinks, and sections.

## How it works

```
VS Code Copilot Chat (agent mode)
        │
        ▼
  powerpoint-dev agent (.md instructions)
        │  describes which PowerShell command to run
        ▼
  PowerPointPOSH module  (imported in the VS Code terminal)
        │  COM calls via PowerPoint Object Model
        ▼
  Microsoft PowerPoint (.pptx / .pptm)
```

- **No separate server** — the module runs directly in the VS Code integrated terminal.
- **No Python / Node** — pure PowerShell 5.1+ on Windows.
- **Full COM access** — everything you can do from VBA, you can do from the agent.
- **-WhatIf / -Confirm** — all state-changing functions support PowerShell's standard risk-mitigation flags.
- **Pester tests** — 20 test files cover every public command group.

## Prerequisites

| Requirement | Details |
|---|---|
| **OS** | Windows 10 / 11 (COM automation is Windows-only) |
| **Microsoft PowerPoint** | PowerPoint 2016, 2019, 2021, or Microsoft 365 (desktop) |
| **PowerShell** | 5.1 (Windows PowerShell) **or** PowerShell 7+ |
| **VS Code** | Latest stable, with the **GitHub Copilot Chat** extension |
| **Copilot** | An active GitHub Copilot subscription |

## Setup

### 1 — Clone the repo

```powershell
git clone https://github.com/jcolozzi/MSPowerPoint-CLI-Agent.git
```

### 2 — Install the agent instructions

Choose **one** of the following:

#### Option A — User-level (available in every workspace)

Copy the `.agent.md` file from the repo root to:
```
C:\Users\%USERNAME%\AppData\Roaming\Code\User\prompts\
```

#### Option B — Workspace-level (scoped to this project)

Copy the `.agent.md` file into a `.github\agents\` folder in your workspace root. VS Code automatically detects any `.md` files in that folder as custom agents.

> [!NOTE]
> VS Code detects any `.md` files in the `.github/agents/` folder of your workspace as custom agents.

### 3 — Update the module path inside the agent file

Open the `.agent.md` file and replace the placeholder path with the actual path to `PowerPointPOSH.psd1` on your machine:

```powershell
# Before
Import-Module "C:\path\to\PowerPointPOSH\PowerPointPOSH.psd1"

# After (example)
Import-Module "C:\Projects\MSPowerPoint-agent\PowerPointPOSH\PowerPointPOSH.psd1"
```

### 4 — Select the agent and start prompting

In VS Code Copilot Chat, click the agent picker and choose **powerpoint-dev**. Open (or have the agent open) a `.pptx` file, then start describing what you want.

## Usage examples

| Prompt | Functions called |
|---|---|
| "Open my presentation deck.pptx" | `Open-PowerPointPresentation` |
| "Add a new slide with a title layout" | `New-PowerPointSlide` |
| "Put a rectangle on slide 2 at position (100, 100)" | `Add-PowerPointShape` |
| "Set the text in Rectangle 1 to 'Hello World'" | `Set-PowerPointText` |
| "Insert a bar chart on slide 3 with this data" | `Add-PowerPointChart` |
| "Add a fade transition to all slides" | `Set-PowerPointTransition` |
| "Export the presentation as PDF" | `Export-PowerPointToPdf` |
| "Show me the VBA code in Module1" | `Get-PowerPointVbaCode` |
| "Add an entrance animation to the title shape" | `Add-PowerPointAnimation` |
| "Export all slides as PNG images" | `Export-PowerPointToImages` |
| "Insert an image from C:\logo.png on slide 1" | `Add-PowerPointImage` |
| "What would happen if I ran New-PowerPointSlide? (dry run)" | `New-PowerPointSlide -WhatIf` |

## Project structure

```text
MSPowerPoint-agent/
├── PowerPointPOSH/              # PowerShell module (the engine)
│   ├── PowerPointPOSH.psd1     # Module manifest (v1.0.0, PS 5.1+, Desktop + Core)
│   ├── PowerPointPOSH.psm1     # Module loader
│   ├── Public/                  # 20 files — one per command category
│   │   ├── ApplicationOps.ps1
│   │   ├── PresentationOps.ps1
│   │   ├── SlideOps.ps1
│   │   ├── ShapeOps.ps1
│   │   └── ...
│   └── Private/                 # Internal helpers (COM session, error formatting, etc.)
├── Tests/                       # Pester test suite — 20 test files
│   ├── PowerPointPOSH.Module.Tests.ps1
│   ├── SlideOps.Tests.ps1
│   └── ...
├── .agent.md                    # Agent instructions (the Copilot Chat prompt)
└── README.md
```

## Running the tests

```powershell
# From the repo root
Invoke-Pester .\Tests\ -Output Detailed
```

> Requires [Pester](https://github.com/pester/Pester) 5.x: `Install-Module Pester -MinimumVersion 5.0 -Force`

## Function reference

<details>
<summary><strong>View all 127 public functions</strong></summary>

| Category | Functions |
|---|---|
| **Application** | `Get-PowerPointApplication`, `Set-PowerPointApplication`, `Get-PowerPointTips` |
| **Presentation** | `Open-PowerPointPresentation`, `New-PowerPointPresentation`, `Save-PowerPointPresentation`, `Close-PowerPointPresentation`, `Get-PowerPointPresentationInfo`, `Copy-PowerPointPresentation`, `Convert-PowerPointPresentation`, `Repair-PowerPointPresentation` |
| **Slide** | `Get-PowerPointSlide`, `New-PowerPointSlide`, `Remove-PowerPointSlide`, `Copy-PowerPointSlide`, `Move-PowerPointSlide`, `Set-PowerPointSlideLayout`, `Get-PowerPointSlideNotes`, `Set-PowerPointSlideNotes`, `Set-PowerPointSlideBackground`, `Get-PowerPointPlaceholder` |
| **Shape** | `Add-PowerPointShape`, `Get-PowerPointShape`, `Remove-PowerPointShape`, `Set-PowerPointShapePosition`, `Set-PowerPointShapeSize`, `Copy-PowerPointShape`, `Group-PowerPointShape`, `Ungroup-PowerPointShape`, `Set-PowerPointShapeZOrder`, `Align-PowerPointShape`, `Distribute-PowerPointShape`, `Rotate-PowerPointShape`, `Lock-PowerPointShape`, `Rename-PowerPointShape` |
| **Text** | `Get-PowerPointText`, `Set-PowerPointText`, `Format-PowerPointText`, `Add-PowerPointTextBox`, `Find-PowerPointText`, `Replace-PowerPointText`, `Set-PowerPointBullet`, `Set-PowerPointParagraphFormat` |
| **Table** | `Add-PowerPointTable`, `Get-PowerPointTableCell`, `Set-PowerPointTableCell`, `Get-PowerPointTable`, `Format-PowerPointTableCell`, `Set-PowerPointTableStyle`, `Add-PowerPointTableRow` |
| **Chart** | `Add-PowerPointChart`, `Get-PowerPointChart`, `Set-PowerPointChartData`, `Set-PowerPointChartStyle`, `Set-PowerPointChartType`, `Export-PowerPointChart` |
| **Image & Media** | `Add-PowerPointImage`, `Add-PowerPointAudio`, `Add-PowerPointVideo`, `Set-PowerPointMediaPlayback`, `Get-PowerPointMedia`, `Export-PowerPointSlideImage`, `Set-PowerPointImageCrop` |
| **Animation** | `Add-PowerPointAnimation`, `Get-PowerPointAnimation`, `Remove-PowerPointAnimation`, `Set-PowerPointAnimationTiming`, `Set-PowerPointAnimationTrigger`, `Set-PowerPointAnimationOrder`, `Copy-PowerPointAnimation` |
| **Transition** | `Set-PowerPointTransition`, `Get-PowerPointTransition`, `Remove-PowerPointTransition`, `Set-PowerPointTransitionTiming` |
| **SlideShow** | `Start-PowerPointSlideShow`, `Stop-PowerPointSlideShow`, `Set-PowerPointSlideShowSettings`, `Invoke-PowerPointSlideShowNavigate`, `Get-PowerPointSlideShowStatus`, `Set-PowerPointSlideShowTimings` |
| **Formatting** | `Set-PowerPointFill`, `Set-PowerPointLine`, `Set-PowerPointShadow`, `Set-PowerPointEffect`, `Set-PowerPointShapeStyle`, `Get-PowerPointFormatting`, `Reset-PowerPointFormatting` |
| **Export** | `Export-PowerPointToPdf`, `Export-PowerPointToImages`, `Export-PowerPointToVideo`, `Export-PowerPointToHtml`, `Export-PowerPointSlide`, `Export-PowerPointHandout` |
| **Master & Layout** | `Get-PowerPointSlideMaster`, `Get-PowerPointSlideLayout`, `New-PowerPointSlideLayout`, `Set-PowerPointSlideMasterElement`, `Copy-PowerPointSlideMaster` |
| **Metadata** | `Get-PowerPointProperty`, `Set-PowerPointProperty`, `Get-PowerPointComment`, `Add-PowerPointComment`, `Remove-PowerPointComment`, `Get-PowerPointTag`, `Set-PowerPointTag` |
| **Print** | `Get-PowerPointPageSetup`, `Set-PowerPointPageSetup`, `Invoke-PowerPointPrint` |
| **VBE** | `Get-PowerPointVbaCode`, `Set-PowerPointVbaCode`, `Add-PowerPointVbaModule`, `Remove-PowerPointVbaModule`, `Invoke-PowerPointVbaMacro`, `Get-PowerPointVbaReference`, `Add-PowerPointVbaReference` |
| **SmartArt** | `Add-PowerPointSmartArt`, `Get-PowerPointSmartArt`, `Set-PowerPointSmartArtNode`, `Set-PowerPointSmartArtLayout` |
| **Hyperlink** | `Add-PowerPointHyperlink`, `Get-PowerPointHyperlink`, `Remove-PowerPointHyperlink` |
| **Section** | `Get-PowerPointSection`, `Add-PowerPointSection`, `Remove-PowerPointSection`, `Rename-PowerPointSection`, `Move-PowerPointSection` |

</details>

All state-changing functions support `-WhatIf` and `-Confirm` via PowerShell's standard `ShouldProcess` mechanism.

## Contributing

Pull requests are welcome. For significant changes, open an issue first to discuss what you would like to change. Please include or update Pester tests for any new or modified functions.

## Credits

- PowerShell port and VS Code agent integration: PowerPointPOSH

## License

[MIT](LICENSE) © 2026 PowerPointPOSH
