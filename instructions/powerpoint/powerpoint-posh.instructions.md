---
description: "Use when the user mentions PowerPoint presentations (.pptx/.pptm/.ppt), PowerPointPOSH, PowerShell COM automation for PowerPoint, or asks to create/edit/format presentations."
---
# PowerPointPOSH Module

When working with Microsoft PowerPoint presentations, import the PowerShell module:

```powershell
Import-Module "K:\Workgrp\PERSONAL SHARE\Colozzi\Access Agent\MSPowerPoint-agent\PowerPointPOSH\PowerPointPOSH.psd1" -Force
```

This module provides 127 PowerShell functions for full PowerPoint presentation automation via COM. Use `-AsJson` on any function for structured output. The `@pptx-dev` agent has the complete function reference.

## Error Handling

If the module fails to import, follow these steps in order:

### Step 1: Verify Module Path
Confirm the path exists and is syntactically correct. Verify the current user has accessible permissions.
- **If this fails:** Proceed to Step 2.

### Step 2: Check File Exists
Verify PowerPointPOSH.psd1 file exists at the specified location.
- **If file is missing:** Check the path with your system administrator.
- **If file exists:** Proceed to Step 3.

### Step 3: Verify Dependencies & Permissions
Check for missing dependencies and verify sufficient permissions are granted. PowerPoint must be installed on the machine.
- **If dependencies are missing:** Install required components (Microsoft Office with PowerPoint).
- **If permissions are insufficient:** Request elevated access.
- **If both are OK:** Proceed to Step 4.

### Step 4: Check Execution Policy
Ensure PowerShell execution policy allows module loading.
- **If policy blocks execution:** Run `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
- **If policy is acceptable:** Proceed to troubleshooting below.

**If Issue Persists**
Consult the system administrator or refer to the PowerPointPOSH documentation for advanced troubleshooting.
