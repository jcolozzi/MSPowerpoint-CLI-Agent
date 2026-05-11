# Public/ChartOps.ps1 — Chart manipulation operations

# ══════════════════════════════════════════════════════════════════════════
# Add-PowerPointChart
# ══════════════════════════════════════════════════════════════════════════

function Add-PowerPointChart {
    <#
    .SYNOPSIS
        Add a chart to a slide.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ChartType
        Chart type key (e.g. 'columnClustered', 'pie', 'line').
    .PARAMETER Left
        Left position in points (default 100).
    .PARAMETER Top
        Top position in points (default 100).
    .PARAMETER Width
        Width in points (default 400).
    .PARAMETER Height
        Height in points (default 300).
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Add-PowerPointChart -SlideIndex 1 -ChartType pie -AsJson
    .EXAMPLE
        Add-PowerPointChart -SlideIndex 2 -ChartType columnClustered -Left 50 -Top 80 -Width 500 -Height 350 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [Parameter(Mandatory)][string]$ChartType,
        [double]$Left   = 100,
        [double]$Top    = 100,
        [double]$Width  = 400,
        [double]$Height = 300,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)

    $chartTypeMap = @{
        'columnClustered' = 51
        'columnStacked'   = 52
        'barClustered'    = 57
        'barStacked'      = 58
        'line'            = 4
        'lineMarkers'     = 65
        'pie'             = 5
        'doughnut'        = -4120
        'area'            = 1
        'areaStacked'     = 76
        'scatter'         = -4169
        'radar'           = -4151
    }

    $chartTypeValue = Resolve-EnumValue -Map $chartTypeMap -Key $ChartType

    if (-not $PSCmdlet.ShouldProcess("Slide $SlideIndex", "Add $ChartType chart")) { return }

    $shape = $slide.Shapes.AddChart2(-1, $chartTypeValue, $Left, $Top, $Width, $Height)

    $result = [ordered]@{
        status    = 'added'
        name      = $shape.Name
        chartType = $ChartType
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Get-PowerPointChart
# ══════════════════════════════════════════════════════════════════════════

function Get-PowerPointChart {
    <#
    .SYNOPSIS
        Get chart information from a shape.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the chart shape.
    .PARAMETER ShapeIndex
        1-based index of the chart shape.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Get-PowerPointChart -SlideIndex 1 -ShapeName "Chart 1" -AsJson
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

    if ($shape.HasChart -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain a chart."),
            'NoChart',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $chart     = $shape.Chart
    $hasTitle  = $chart.HasTitle
    $titleText = if ($hasTitle) { $chart.ChartTitle.Text } else { $null }
    $hasLegend = $chart.HasLegend
    $seriesCount = $chart.SeriesCollection().Count

    $result = [ordered]@{
        shapeName   = $shape.Name
        chartType   = [int]$chart.ChartType
        hasTitle    = [bool]$hasTitle
        title       = $titleText
        hasLegend   = [bool]$hasLegend
        seriesCount = $seriesCount
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointChart
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointChart {
    <#
    .SYNOPSIS
        Modify chart properties (title, legend, type).
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the chart shape.
    .PARAMETER ShapeIndex
        1-based index of the chart shape.
    .PARAMETER Title
        Chart title text. Sets HasTitle to true.
    .PARAMETER HasLegend
        Show or hide the legend.
    .PARAMETER ChartType
        Change chart type (e.g. 'pie', 'line').
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointChart -SlideIndex 1 -ShapeName "Chart 1" -Title "Sales Q1" -HasLegend $true -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [string]$Title,
        [System.Nullable[bool]]$HasLegend,
        [string]$ChartType,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasChart -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain a chart."),
            'NoChart',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Chart '$target' on slide $SlideIndex", "Set chart properties")) { return }

    $chart   = $shape.Chart
    $changes = @()

    if ($PSBoundParameters.ContainsKey('Title')) {
        $chart.HasTitle = $true
        $chart.ChartTitle.Text = $Title
        $changes += 'title'
    }
    if ($PSBoundParameters.ContainsKey('HasLegend')) {
        $chart.HasLegend = $HasLegend
        $changes += 'hasLegend'
    }
    if ($PSBoundParameters.ContainsKey('ChartType')) {
        $chartTypeMap = @{
            'columnClustered' = 51
            'columnStacked'   = 52
            'barClustered'    = 57
            'barStacked'      = 58
            'line'            = 4
            'lineMarkers'     = 65
            'pie'             = 5
            'doughnut'        = -4120
            'area'            = 1
            'areaStacked'     = 76
            'scatter'         = -4169
            'radar'           = -4151
        }
        $chartTypeValue = Resolve-EnumValue -Map $chartTypeMap -Key $ChartType
        $chart.ChartType = $chartTypeValue
        $changes += 'chartType'
    }

    $result = [ordered]@{
        status  = 'updated'
        shape   = $shape.Name
        changed = $changes
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointChartData
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointChartData {
    <#
    .SYNOPSIS
        Set chart data (categories and series values) via the embedded Excel workbook.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the chart shape.
    .PARAMETER ShapeIndex
        1-based index of the chart shape.
    .PARAMETER Categories
        Array of category labels.
    .PARAMETER SeriesName
        Name of the data series.
    .PARAMETER Values
        Array of numeric values corresponding to each category.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointChartData -SlideIndex 1 -ShapeName "Chart 1" -Categories @("Q1","Q2","Q3") -SeriesName "Revenue" -Values @(100,200,150) -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][string[]]$Categories,
        [Parameter(Mandatory)][string]$SeriesName,
        [Parameter(Mandatory)][double[]]$Values,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasChart -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain a chart."),
            'NoChart',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Chart '$target' on slide $SlideIndex", "Set chart data")) { return }

    $chart = $shape.Chart
    $chart.ChartData.Activate()
    $wb = $chart.ChartData.Workbook
    $ws = $wb.Worksheets(1)

    try {
        # Write series header
        $ws.Cells(1, 2).Value2 = $SeriesName

        # Write categories and values
        for ($i = 0; $i -lt $Categories.Count; $i++) {
            $ws.Cells($i + 2, 1).Value2 = $Categories[$i]
            $ws.Cells($i + 2, 2).Value2 = $Values[$i]
        }
    } finally {
        $wb.Close($false)
    }

    $result = [ordered]@{
        status     = 'updated'
        shape      = $shape.Name
        seriesName = $SeriesName
        categories = $Categories.Count
        values     = $Values.Count
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Set-PowerPointChartSeries
# ══════════════════════════════════════════════════════════════════════════

function Set-PowerPointChartSeries {
    <#
    .SYNOPSIS
        Format a chart series (color, line weight).
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the chart shape.
    .PARAMETER ShapeIndex
        1-based index of the chart shape.
    .PARAMETER SeriesIndex
        1-based index of the data series.
    .PARAMETER Color
        Series color as 'R,G,B' string.
    .PARAMETER LineWeight
        Line weight in points.
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Set-PowerPointChartSeries -SlideIndex 1 -ShapeName "Chart 1" -SeriesIndex 1 -Color "255,0,0" -LineWeight 2.5 -AsJson
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][int]$SeriesIndex,
        [string]$Color,
        [double]$LineWeight,
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasChart -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain a chart."),
            'NoChart',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $target = if ($ShapeName) { $ShapeName } else { "index $ShapeIndex" }
    if (-not $PSCmdlet.ShouldProcess("Series $SeriesIndex in chart '$target' on slide $SlideIndex", "Format series")) { return }

    $series  = $shape.Chart.SeriesCollection($SeriesIndex)
    $changes = @()

    if ($PSBoundParameters.ContainsKey('Color')) {
        $series.Format.Fill.Solid()
        $series.Format.Fill.ForeColor.RGB = ConvertFrom-RGBString $Color
        $changes += 'color'
    }
    if ($PSBoundParameters.ContainsKey('LineWeight')) {
        $series.Format.Line.Weight = $LineWeight
        $changes += 'lineWeight'
    }

    $result = [ordered]@{
        status      = 'formatted'
        shape       = $shape.Name
        seriesIndex = $SeriesIndex
        changed     = $changes
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}

# ══════════════════════════════════════════════════════════════════════════
# Export-PowerPointChart
# ══════════════════════════════════════════════════════════════════════════

function Export-PowerPointChart {
    <#
    .SYNOPSIS
        Export a chart to an image file.
    .PARAMETER PresentationPath
        Path to the presentation file. Uses current session if omitted.
    .PARAMETER SlideIndex
        1-based slide index.
    .PARAMETER ShapeName
        Name of the chart shape.
    .PARAMETER ShapeIndex
        1-based index of the chart shape.
    .PARAMETER FilePath
        Output file path for the exported image.
    .PARAMETER Format
        Image format: 'png', 'jpg', 'gif', or 'bmp' (default 'png').
    .PARAMETER AsJson
        Return JSON string instead of PSCustomObject.
    .EXAMPLE
        Export-PowerPointChart -SlideIndex 1 -ShapeName "Chart 1" -FilePath "C:\chart.png" -AsJson
    #>
    [CmdletBinding()]
    param(
        [string]$PresentationPath,
        [Parameter(Mandatory)][int]$SlideIndex,
        [string]$ShapeName,
        [int]$ShapeIndex,
        [Parameter(Mandatory)][string]$FilePath,
        [ValidateSet('png','jpg','gif','bmp')][string]$Format = 'png',
        [switch]$AsJson
    )

    $resolvedPath = Resolve-SessionPresentationPath -PresentationPath $PresentationPath -CallerName $MyInvocation.MyCommand.Name
    $app   = Connect-PowerPointPresentation -PresentationPath $resolvedPath
    $pres  = $app.ActivePresentation
    $slide = $pres.Slides.Item($SlideIndex)
    $shape = Get-SlideShape -Slide $slide -ShapeName $ShapeName -ShapeIndex $ShapeIndex

    if ($shape.HasChart -ne -1) {
        $er = [System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new("Shape '$($shape.Name)' does not contain a chart."),
            'NoChart',
            [System.Management.Automation.ErrorCategory]::InvalidOperation, $shape.Name)
        $PSCmdlet.ThrowTerminatingError($er)
    }

    $resolvedFile = [System.IO.Path]::GetFullPath($FilePath)
    $shape.Chart.Export($resolvedFile, $Format)

    $result = [ordered]@{
        status   = 'exported'
        shape    = $shape.Name
        filePath = $resolvedFile
        format   = $Format
    }
    Format-PowerPointOutput -Data $result -AsJson:$AsJson
}
