# Public/TableOps.ps1 — Table manipulation operations

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointTable
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointTable {
    <#
    .SYNOPSIS
        Add a table to a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER Rows
        Number of rows.
    .PARAMETER Columns
        Number of columns.
    .PARAMETER Left
        Left position in points (default 100).
    .PARAMETER Top
        Top position in points (default 100).
    .PARAMETER Width
        Width in points (default 500).
    .PARAMETER Height
        Height in points (default 300).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointTable -SlideIndex 1 -Rows 3 -Columns 4 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][int]$Rows,
        [Parameter(Mandatory)][int]$Columns,
        [double]$Left   = 100,
        [double]$Top    = 100,
        [double]$Width  = 500,
        [double]$Height = 300,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex", "Add ${Rows}x${Columns} table")) { return }

    $shape = $slide.Shapes.AddTable($Rows, $Columns, $Left, $Top, $Width, $Height)

    $result = [ordered]@{
        status  = 'added'
        name    = $shape.Name
        index   = $shape.ZOrderPosition
        rows    = $Rows
        columns = $Columns
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointTable
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointTable {
    <#
    .SYNOPSIS
        Get table data from a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the table shape.
    .PARAMETER ShapeIndex
        1-based index of the table shape.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointTable -SlideIndex 1 -ShapeName "Table 1" -AsJson
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

    if ($shape.HasTable -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain a table."),
            'NoTable',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $table    = $shape.Table
    $rowCount = $table.Rows.Count
    $colCount = $table.Columns.Count
    $data     = @()

    for ($r = 1; $r -le $rowCount; $r++) {
        $rowData = @()
        for ($c = 1; $c -le $colCount; $c++) {
            $rowData += $table.Cell($r, $c).Shape.TextFrame.TextRange.Text
        }
        $data += , $rowData
    }

    $result = [ordered]@{
        shapeName = $shape.Name
        rows      = $rowCount
        columns   = $colCount
        data      = $data
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointTableCell
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointTableCell {
    <#
    .SYNOPSIS
        Set the text value of a table cell.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the table shape.
    .PARAMETER ShapeIndex
        1-based index of the table shape.
    .PARAMETER Row
        1-based row number.
    .PARAMETER Column
        1-based column number.
    .PARAMETER Value
        Text value to set.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointTableCell -SlideIndex 1 -ShapeName "Table 1" -Row 1 -Column 1 -Value "Header" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][int]$Row,
        [Parameter(Mandatory)][int]$Column,
        [Parameter(Mandatory)][string]$Value,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasTable -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain a table."),
            'NoTable',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Cell ($Row,$Column) in '$target' on slide $SlideIndex", "Set value")) { return }

    $shape.Table.Cell($Row, $Column).Shape.TextFrame.TextRange.Text = $Value

    $result = [ordered]@{
        status = 'updated'
        shape  = $shape.Name
        row    = $Row
        column = $Column
        value  = $Value
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointTableCell
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointTableCell {
    <#
    .SYNOPSIS
        Get the text value of a table cell.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the table shape.
    .PARAMETER ShapeIndex
        1-based index of the table shape.
    .PARAMETER Row
        1-based row number.
    .PARAMETER Column
        1-based column number.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointTableCell -SlideIndex 1 -ShapeName "Table 1" -Row 2 -Column 3 -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][int]$Row,
        [Parameter(Mandatory)][int]$Column,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasTable -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain a table."),
            'NoTable',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $text = $shape.Table.Cell($Row, $Column).Shape.TextFrame.TextRange.Text

    $result = [ordered]@{
        shape  = $shape.Name
        row    = $Row
        column = $Column
        value  = $text
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Format-PowerPointTableCell
# ══════════════════════════════════════════════════════════════════════════

function Format-PowerPointTableCell {
    <#
    .SYNOPSIS
        Format a table cell (font, fill).
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the table shape.
    .PARAMETER ShapeIndex
        1-based index of the table shape.
    .PARAMETER Row
        1-based row number.
    .PARAMETER Column
        1-based column number.
    .PARAMETER FontName
        Font family name.
    .PARAMETER FontSize
        Font size in points.
    .PARAMETER Bold
        Set bold formatting.
    .PARAMETER Italic
        Set italic formatting.
    .PARAMETER FillColor
        Cell background color as 'R,G,B' string (e.g. '255,200,0').
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Format-PowerPointTableCell -SlideIndex 1 -ShapeName "Table 1" -Row 1 -Column 1 -Bold $true -FillColor "0,120,215" -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][int]$Row,
        [Parameter(Mandatory)][int]$Column,
        [string]$FontName,
        [double]$FontSize,
        [System.Nullable[bool]]$Bold,
        [System.Nullable[bool]]$Italic,
        [string]$FillColor,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasTable -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain a table."),
            'NoTable',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Cell ($Row,$Column) in '$target' on slide $SlideIndex", "Format cell")) { return }

    $cell = $shape.Table.Cell($Row, $Column)
    $font = $cell.Shape.TextFrame.TextRange.Font

    $changes = @()
    if ($PSBoundParameters.ContainsKey('FontName'))  { $font.Name   = $FontName;  $changes += 'fontName' }
    if ($PSBoundParameters.ContainsKey('FontSize'))  { $font.Size   = $FontSize;  $changes += 'fontSize' }
    if ($PSBoundParameters.ContainsKey('Bold'))       { $font.Bold   = if ($Bold)   { -1 } else { 0 }; $changes += 'bold' }
    if ($PSBoundParameters.ContainsKey('Italic'))     { $font.Italic = if ($Italic) { -1 } else { 0 }; $changes += 'italic' }
    if ($PSBoundParameters.ContainsKey('FillColor')) {
        $cell.Shape.Fill.Solid()
        $cell.Shape.Fill.ForeColor.RGB = ConvertFrom-RGBString $FillColor
        $changes += 'fillColor'
    }

    $result = [ordered]@{
        status  = 'formatted'
        shape   = $shape.Name
        row     = $Row
        column  = $Column
        changed = $changes
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointTableRow
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointTableRow {
    <#
    .SYNOPSIS
        Add a row to a table.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the table shape.
    .PARAMETER ShapeIndex
        1-based index of the table shape.
    .PARAMETER BeforeRow
        1-based row index to insert before. Omit to append at the end.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointTableRow -SlideIndex 1 -ShapeName "Table 1" -AsJson
    .EXAMPLE
        Add-PowerPointTableRow -SlideIndex 1 -ShapeName "Table 1" -BeforeRow 2 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [int]$BeforeRow,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasTable -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain a table."),
            'NoTable',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    $action = if ($BeforeRow) { "Insert row before row $BeforeRow" } else { "Append row" }
    if (-not $PSCmdlet.ShouldProcess("Table '$target' on slide $SlideIndex", $action)) { return }

    $table = $shape.Table
    if ($PSBoundParameters.ContainsKey('BeforeRow')) {
        $table.Rows.Add($BeforeRow)
    } else {
        $table.Rows.Add()
    }

    $result = [ordered]@{
        status   = 'added'
        shape    = $shape.Name
        rowCount = $table.Rows.Count
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Remove-PowerPointTableRow
# ══════════════════════════════════════════════════════════════════════════

function Remove-PowerPointTableRow {
    <#
    .SYNOPSIS
        Remove a row from a table.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the table shape.
    .PARAMETER ShapeIndex
        1-based index of the table shape.
    .PARAMETER Row
        1-based row index to remove.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Remove-PowerPointTableRow -SlideIndex 1 -ShapeName "Table 1" -Row 3 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][int]$Row,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasTable -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain a table."),
            'NoTable',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Row $Row in table '$target' on slide $SlideIndex", "Remove row")) { return }

    $table = $shape.Table
    $table.Rows.Item($Row).Delete()

    $result = [ordered]@{
        status   = 'removed'
        shape    = $shape.Name
        rowCount = $table.Rows.Count
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
