function Get-PowerPointVbaCode {
    <#
    .SYNOPSIS
        Reads VBA code from a specific module in the active PowerPoint presentation.
    .PARAMETER PresentationPath
        Path to a .pptm presentation file. If omitted, uses the current session presentation.
    .PARAMETER ModuleName
        Name of the VBA module to read.
    .PARAMETER AsJson
        Return output as a JSON string.
    .EXAMPLE
        Get-PowerPointVbaCode -ModuleName 'Module1'
    .EXAMPLE
        Get-PowerPointVbaCode -PresentationPath 'C:\Decks\Report.pptm' -ModuleName 'Utils' -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$PresentationPath,

        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter()]
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointVbaCode'
    $ext = [System.IO.Path]::GetExtension($PresentationPath).ToLower()
    if ($ext -notin '.pptm', '.ppt', '.ppa', '.ppam') {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBE operations require a macro-enabled file (.pptm). Current file: $ext"),
            'VBE_NotMacroEnabled', [System.Management.Automation.ErrorCategory]::InvalidOperation, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $app = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    try {
        $vbProj = $pres.VBProject
    }
    catch {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBA project access denied. Enable 'Trust access to the VBA project object model' in PowerPoint Trust Center settings."),
            'VBE_AccessDenied', [System.Management.Automation.ErrorCategory]::PermissionDenied, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $mod = $vbProj.VBComponents.Item($ModuleName).CodeModule
    $lineCount = $mod.CountOfLines
    $code = if ($lineCount -gt 0) { $mod.Lines(1, $lineCount) } else { '' }

    $result = [ordered]@{
        status     = 'success'
        moduleName = $ModuleName
        lineCount  = $lineCount
        code       = $code
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Set-PowerPointVbaCode {
    <#
    .SYNOPSIS
        Replaces VBA code in a specific module in the active PowerPoint presentation.
    .PARAMETER PresentationPath
        Path to a .pptm presentation file. If omitted, uses the current session presentation.
    .PARAMETER ModuleName
        Name of the VBA module to write to.
    .PARAMETER Code
        The full VBA code to set in the module.
    .PARAMETER AsJson
        Return output as a JSON string.
    .EXAMPLE
        Set-PowerPointVbaCode -ModuleName 'Module1' -Code 'Sub Hello()\n    MsgBox "Hi"\nEnd Sub'
    .EXAMPLE
        Set-PowerPointVbaCode -PresentationPath 'C:\Decks\Report.pptm' -ModuleName 'Utils' -Code $vbaSource
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string]$PresentationPath,

        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter(Mandatory)]
        [string]$Code,

        [Parameter()]
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Set-PowerPointVbaCode'
    $ext = [System.IO.Path]::GetExtension($PresentationPath).ToLower()
    if ($ext -notin '.pptm', '.ppt', '.ppa', '.ppam') {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBE operations require a macro-enabled file (.pptm). Current file: $ext"),
            'VBE_NotMacroEnabled', [System.Management.Automation.ErrorCategory]::InvalidOperation, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $app = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    try {
        $vbProj = $pres.VBProject
    }
    catch {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBA project access denied. Enable 'Trust access to the VBA project object model' in PowerPoint Trust Center settings."),
            'VBE_AccessDenied', [System.Management.Automation.ErrorCategory]::PermissionDenied, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    if ($PSCmdlet.ShouldProcess("Module '$ModuleName'", 'Replace VBA code')) {
        $mod = $vbProj.VBComponents.Item($ModuleName).CodeModule
        if ($mod.CountOfLines -gt 0) {
            $mod.DeleteLines(1, $mod.CountOfLines)
        }
        $mod.AddFromString($Code)

        $result = [ordered]@{
            status     = 'success'
            moduleName = $ModuleName
            lineCount  = $mod.CountOfLines
        }
        Format-PowerPointOutput -Data $result -AsJson:$AsJson
    }
}

function Add-PowerPointVbaModule {
    <#
    .SYNOPSIS
        Adds a new VBA module to the active PowerPoint presentation.
    .PARAMETER PresentationPath
        Path to a .pptm presentation file. If omitted, uses the current session presentation.
    .PARAMETER ModuleName
        Name for the new VBA module.
    .PARAMETER ModuleType
        Type of module to create: 'standard' or 'class'. Defaults to 'standard'.
    .PARAMETER AsJson
        Return output as a JSON string.
    .EXAMPLE
        Add-PowerPointVbaModule -ModuleName 'UtilityHelpers'
    .EXAMPLE
        Add-PowerPointVbaModule -ModuleName 'clsWidget' -ModuleType 'class'
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string]$PresentationPath,

        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter()]
        [ValidateSet('standard', 'class')]
        [string]$ModuleType = 'standard',

        [Parameter()]
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Add-PowerPointVbaModule'
    $ext = [System.IO.Path]::GetExtension($PresentationPath).ToLower()
    if ($ext -notin '.pptm', '.ppt', '.ppa', '.ppam') {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBE operations require a macro-enabled file (.pptm). Current file: $ext"),
            'VBE_NotMacroEnabled', [System.Management.Automation.ErrorCategory]::InvalidOperation, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $app = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    try {
        $vbProj = $pres.VBProject
    }
    catch {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBA project access denied. Enable 'Trust access to the VBA project object model' in PowerPoint Trust Center settings."),
            'VBE_AccessDenied', [System.Management.Automation.ErrorCategory]::PermissionDenied, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    # vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2
    $typeValue = if ($ModuleType -eq 'class') { 2 } else { 1 }

    if ($PSCmdlet.ShouldProcess("Module '$ModuleName' ($ModuleType)", 'Add VBA module')) {
        $comp = $vbProj.VBComponents.Add($typeValue)
        $comp.Name = $ModuleName

        $result = [ordered]@{
            status     = 'success'
            moduleName = $ModuleName
            moduleType = $ModuleType
        }
        Format-PowerPointOutput -Data $result -AsJson:$AsJson
    }
}

function Remove-PowerPointVbaModule {
    <#
    .SYNOPSIS
        Removes a VBA module from the active PowerPoint presentation.
    .PARAMETER PresentationPath
        Path to a .pptm presentation file. If omitted, uses the current session presentation.
    .PARAMETER ModuleName
        Name of the VBA module to remove.
    .PARAMETER AsJson
        Return output as a JSON string.
    .EXAMPLE
        Remove-PowerPointVbaModule -ModuleName 'Module1'
    .EXAMPLE
        Remove-PowerPointVbaModule -PresentationPath 'C:\Decks\Report.pptm' -ModuleName 'OldUtils' -Confirm:$false
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string]$PresentationPath,

        [Parameter(Mandatory)]
        [string]$ModuleName,

        [Parameter()]
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Remove-PowerPointVbaModule'
    $ext = [System.IO.Path]::GetExtension($PresentationPath).ToLower()
    if ($ext -notin '.pptm', '.ppt', '.ppa', '.ppam') {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBE operations require a macro-enabled file (.pptm). Current file: $ext"),
            'VBE_NotMacroEnabled', [System.Management.Automation.ErrorCategory]::InvalidOperation, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $app = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    try {
        $vbProj = $pres.VBProject
    }
    catch {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBA project access denied. Enable 'Trust access to the VBA project object model' in PowerPoint Trust Center settings."),
            'VBE_AccessDenied', [System.Management.Automation.ErrorCategory]::PermissionDenied, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    if ($PSCmdlet.ShouldProcess("Module '$ModuleName'", 'Remove VBA module')) {
        $comp = $vbProj.VBComponents.Item($ModuleName)
        $vbProj.VBComponents.Remove($comp)

        $result = [ordered]@{
            status     = 'success'
            moduleName = $ModuleName
        }
        Format-PowerPointOutput -Data $result -AsJson:$AsJson
    }
}

function Get-PowerPointVbaModuleList {
    <#
    .SYNOPSIS
        Lists all VBA modules in the active PowerPoint presentation.
    .PARAMETER PresentationPath
        Path to a .pptm presentation file. If omitted, uses the current session presentation.
    .PARAMETER AsJson
        Return output as a JSON string.
    .EXAMPLE
        Get-PowerPointVbaModuleList
    .EXAMPLE
        Get-PowerPointVbaModuleList -PresentationPath 'C:\Decks\Report.pptm' -AsJson
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$PresentationPath,

        [Parameter()]
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointVbaModuleList'
    $ext = [System.IO.Path]::GetExtension($PresentationPath).ToLower()
    if ($ext -notin '.pptm', '.ppt', '.ppa', '.ppam') {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBE operations require a macro-enabled file (.pptm). Current file: $ext"),
            'VBE_NotMacroEnabled', [System.Management.Automation.ErrorCategory]::InvalidOperation, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $app = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    try {
        $vbProj = $pres.VBProject
    }
    catch {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBA project access denied. Enable 'Trust access to the VBA project object model' in PowerPoint Trust Center settings."),
            'VBE_AccessDenied', [System.Management.Automation.ErrorCategory]::PermissionDenied, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $typeMap = @{ 1 = 'standard'; 2 = 'class'; 3 = 'form'; 100 = 'document' }
    $modules = [System.Collections.Generic.List[object]]::new()

    foreach ($comp in $vbProj.VBComponents) {
        $typeNum = [int]$comp.Type
        $typeName = if ($typeMap.ContainsKey($typeNum)) { $typeMap[$typeNum] } else { "unknown($typeNum)" }
        $modules.Add([ordered]@{
            name      = $comp.Name
            type      = $typeName
            lineCount = $comp.CodeModule.CountOfLines
        })
    }

    $result = [ordered]@{
        status  = 'success'
        modules = @($modules)
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Find-PowerPointVbaCode {
    <#
    .SYNOPSIS
        Searches for text across all VBA modules in the active PowerPoint presentation.
    .PARAMETER PresentationPath
        Path to a .pptm presentation file. If omitted, uses the current session presentation.
    .PARAMETER SearchText
        The text string to search for in VBA code.
    .PARAMETER MatchCase
        Perform a case-sensitive search. Defaults to $false.
    .PARAMETER AsJson
        Return output as a JSON string.
    .EXAMPLE
        Find-PowerPointVbaCode -SearchText 'MsgBox'
    .EXAMPLE
        Find-PowerPointVbaCode -SearchText 'Option Explicit' -MatchCase
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$PresentationPath,

        [Parameter(Mandatory)]
        [string]$SearchText,

        [Parameter()]
        [switch]$MatchCase,

        [Parameter()]
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Find-PowerPointVbaCode'
    $ext = [System.IO.Path]::GetExtension($PresentationPath).ToLower()
    if ($ext -notin '.pptm', '.ppt', '.ppa', '.ppam') {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBE operations require a macro-enabled file (.pptm). Current file: $ext"),
            'VBE_NotMacroEnabled', [System.Management.Automation.ErrorCategory]::InvalidOperation, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $app = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    try {
        $vbProj = $pres.VBProject
    }
    catch {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBA project access denied. Enable 'Trust access to the VBA project object model' in PowerPoint Trust Center settings."),
            'VBE_AccessDenied', [System.Management.Automation.ErrorCategory]::PermissionDenied, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $comparison = if ($MatchCase) {
        [System.StringComparison]::Ordinal
    }
    else {
        [System.StringComparison]::OrdinalIgnoreCase
    }

    $matches = [System.Collections.Generic.List[object]]::new()

    foreach ($comp in $vbProj.VBComponents) {
        $mod = $comp.CodeModule
        $totalLines = $mod.CountOfLines
        for ($i = 1; $i -le $totalLines; $i++) {
            $lineText = $mod.Lines($i, 1)
            if ($lineText.IndexOf($SearchText, $comparison) -ge 0) {
                $matches.Add([ordered]@{
                    moduleName = $comp.Name
                    lineNumber = $i
                    lineText   = $lineText
                })
            }
        }
    }

    $result = [ordered]@{
        status     = 'success'
        searchText = $SearchText
        matchCount = $matches.Count
        matches    = @($matches)
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Invoke-PowerPointVbaMacro {
    <#
    .SYNOPSIS
        Runs a VBA macro (Sub or Function) in the active PowerPoint presentation.
    .PARAMETER PresentationPath
        Path to a .pptm presentation file. If omitted, uses the current session presentation.
    .PARAMETER MacroName
        Fully qualified macro name, e.g. 'Module1.MyMacro'.
    .PARAMETER Arguments
        Optional array of arguments to pass to the macro (up to 30).
    .PARAMETER AsJson
        Return output as a JSON string.
    .EXAMPLE
        Invoke-PowerPointVbaMacro -MacroName 'Module1.FormatAllSlides'
    .EXAMPLE
        Invoke-PowerPointVbaMacro -MacroName 'Utils.CalcTotal' -Arguments @(100, 0.08)
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string]$PresentationPath,

        [Parameter(Mandatory)]
        [string]$MacroName,

        [Parameter()]
        [object[]]$Arguments,

        [Parameter()]
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Invoke-PowerPointVbaMacro'
    $ext = [System.IO.Path]::GetExtension($PresentationPath).ToLower()
    if ($ext -notin '.pptm', '.ppt', '.ppa', '.ppam') {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBE operations require a macro-enabled file (.pptm). Current file: $ext"),
            'VBE_NotMacroEnabled', [System.Management.Automation.ErrorCategory]::InvalidOperation, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $app = Connect-PowerPointPresentation -PresentationPath $PresentationPath
    $pres = $app.ActivePresentation

    # Verify VBE access is available
    try {
        $null = $pres.VBProject
    }
    catch {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("VBA project access denied. Enable 'Trust access to the VBA project object model' in PowerPoint Trust Center settings."),
            'VBE_AccessDenied', [System.Management.Automation.ErrorCategory]::PermissionDenied, $PresentationPath)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    if ($PSCmdlet.ShouldProcess("Macro '$MacroName'", 'Run VBA macro')) {
        $macroResult = $null

        if (-not $Arguments -or $Arguments.Count -eq 0) {
            $macroResult = $app.Run($MacroName)
        }
        else {
            if ($Arguments.Count -gt 30) {
                $er = [System.Management.Automation.ErrorRecord]::new(
                    [System.InvalidOperationException]::new("Application.Run supports a maximum of 30 arguments. Received: $($Arguments.Count)."),
                    'VBE_TooManyArgs', [System.Management.Automation.ErrorCategory]::InvalidArgument, $Arguments)
                $PSCmdlet.ThrowTerminatingError($er)
            }

            # Build positional argument list for Application.Run
            $runArgs = [System.Collections.Generic.List[object]]::new()
            $runArgs.Add($MacroName)
            foreach ($a in $Arguments) { $runArgs.Add($a) }

            # Pad with [System.Type]::Missing up to 31 total params (name + 30 args)
            while ($runArgs.Count -lt 31) {
                $runArgs.Add([System.Type]::Missing)
            }

            $macroResult = $app.GetType().InvokeMember(
                'Run',
                [System.Reflection.BindingFlags]::InvokeMethod,
                $null,
                $app,
                $runArgs.ToArray()
            )
        }

        $result = [ordered]@{
            status    = 'success'
            macroName = $MacroName
            result    = $macroResult
        }
        Format-PowerPointOutput -Data $result -AsJson:$AsJson
    }
}
