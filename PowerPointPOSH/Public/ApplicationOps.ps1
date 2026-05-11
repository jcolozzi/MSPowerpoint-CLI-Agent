# Public/ApplicationOps.ps1 — Application-level: info, options, tips

function Get-PowerPointApplicationInfo {
    <#
    .SYNOPSIS
        Return PowerPoint application-level information.
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointApplicationInfo -AsJson
    .EXAMPLE
        Get-PowerPointApplicationInfo -PresentationPath "C:\deck.pptx"
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Get-PowerPointApplicationInfo'
    $app = Connect-PowerPointPresentation -PresentationPath $PresentationPath

    $activePres = $null
    try {
        $activePres = $app.ActivePresentation.Name
    } catch {
        $activePres = $null
    }

    $hwnd = $null
    try { $hwnd = $app.HWND } catch {}

    $windowState = $null
    try {
        $wsVal = $app.ActiveWindow.WindowState
        foreach ($entry in $script:PPT_WINDOW_STATE.GetEnumerator()) {
            if ($entry.Value -eq $wsVal) { $windowState = $entry.Key; break }
        }
        if ($null -eq $windowState) { $windowState = $wsVal }
    } catch {}

    $result = [ordered]@{
        status             = 'ok'
        version            = $app.Version
        name               = $app.Name
        path               = $app.Path
        hwnd               = $hwnd
        operatingSystem    = $app.OperatingSystem
        activePresentation = $activePres
        windowState        = $windowState
        displayAlerts      = $app.DisplayAlerts
        visible            = [bool]($app.Visible)
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Set-PowerPointOption {
    <#
    .SYNOPSIS
        Set PowerPoint application-level options (display alerts, window state, visibility).
    .PARAMETER PresentationPath
        Path to presentation. Falls back to current session presentation.
    .PARAMETER DisplayAlerts
        Alert level: 'none' or 'all'.
    .PARAMETER WindowState
        Window state: 'normal', 'minimized', or 'maximized'.
    .PARAMETER Visible
        Whether the application window is visible.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointOption -DisplayAlerts 'none' -AsJson
    .EXAMPLE
        Set-PowerPointOption -WindowState 'maximized' -Visible $true
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [ValidateSet('none','all')]
        [string]$DisplayAlerts,
        [ValidateSet('normal','minimized','maximized')]
        [string]$WindowState,
        [System.Nullable[bool]]$Visible,
        [switch]$AsJson
    )

    $PresentationPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName 'Set-PowerPointOption'
    $app = Connect-PowerPointPresentation -PresentationPath $PresentationPath

    $changes = @()

    if ($PSBoundParameters.ContainsKey('DisplayAlerts')) {
        $alertVal = Resolve-EnumValue -Map $script:PPT_ALERT_LEVEL -Key $DisplayAlerts
        if ($PSCmdlet.ShouldProcess('PowerPoint Application', "Set DisplayAlerts to '$DisplayAlerts' ($alertVal)")) {
            $app.DisplayAlerts = $alertVal
            $changes += "DisplayAlerts=$DisplayAlerts"
        }
    }

    if ($PSBoundParameters.ContainsKey('WindowState')) {
        $wsVal = Resolve-EnumValue -Map $script:PPT_WINDOW_STATE -Key $WindowState
        if ($PSCmdlet.ShouldProcess('PowerPoint Application', "Set WindowState to '$WindowState' ($wsVal)")) {
            $app.ActiveWindow.WindowState = $wsVal
            $changes += "WindowState=$WindowState"
        }
    }

    if ($PSBoundParameters.ContainsKey('Visible')) {
        if ($PSCmdlet.ShouldProcess('PowerPoint Application', "Set Visible to $Visible")) {
            $visVal = if ($Visible) { -1 } else { 0 }  # msoTrue / msoFalse
            $app.Visible = $visVal
            $changes += "Visible=$Visible"
        }
    }

    $result = [ordered]@{
        status  = 'ok'
        changes = $changes
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

function Get-PowerPointTip {
    <#
    .SYNOPSIS
        Return a random usage tip about the PowerPointPOSH module.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointTip
    .EXAMPLE
        Get-PowerPointTip -AsJson
    #>
    [CmdletBinding()]
    param(
        [switch]$AsJson
    )

    $tips = @(
        'Use -AsJson on any command to get JSON output for AI agent consumption.'
        'Close-PowerPointPresentation -QuitApp releases the COM object and frees the file lock.'
        'Pipe Get-PowerPointSlide output to explore all slides in a presentation.'
        'Set-PowerPointSlideBackground supports both RGB ("255,0,0") and hex ("#FF0000") color formats.'
        'Use Get-PowerPointSlidePlaceholders to discover placeholder indices before setting text.'
        'Convert-PowerPointPresentation can export to PDF, PNG, MP4, and many other formats.'
        'New-PowerPointSlide -Layout "blank" creates a clean slide; use "title" for a title slide.'
        'Copy-PowerPointSlide duplicates a slide; Move-PowerPointSlide reorders without duplicating.'
        'Get-PowerPointSlideNotes reads speaker notes; Set-PowerPointSlideNotes writes them.'
        'Repair-PowerPointPresentation opens, force-saves, and closes — useful for recovering corrupt files.'
        'Get-PowerPointApplicationInfo shows version, HWND, and current window state.'
        'Set-PowerPointOption -WindowState "maximized" maximizes the PowerPoint window.'
        'Always call Close-PowerPointPresentation when done to release COM resources.'
        'New-PowerPointPresentation creates a blank deck and saves it to disk immediately.'
        'Save-PowerPointPresentation returns the last-saved timestamp in ISO 8601 format.'
    )

    $tip = $tips | Get-Random

    $result = [ordered]@{
        status = 'ok'
        tip    = $tip
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
