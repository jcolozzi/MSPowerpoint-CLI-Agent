# Private/Utilities.ps1 — Value conversion, output formatting, and enum resolution

function ConvertTo-PowerPointSafeValue {
    <#
    .SYNOPSIS
        Convert COM values to PowerShell-safe types for JSON serialization.
    #>
    param([AllowNull()]$Value)

    if ($null -eq $Value)               { return $null }
    if ($Value -is [System.DBNull])     { return $null }
    if ($Value -is [System.DateTime])   { return $Value.ToString('o') }  # ISO 8601
    if ($Value -is [decimal])           { return [double]$Value }
    if ($Value -is [byte[]])            { return "<binary $($Value.Length) bytes>" }
    # Handle COM error values (e.g. #N/A, #VALUE!)
    if ($Value -is [int] -and $Value -eq -2146826246) { return '#N/A' }
    return $Value
}

function Format-PowerPointOutput {
    <#
    .SYNOPSIS
        Handle -AsJson switch: convert hashtable/PSCustomObject to JSON or return as-is.
    #>
    param(
        [Parameter(Mandatory)]$Data,
        [switch]$AsJson
    )

    if ($Data -is [hashtable]) {
        $Data = [PSCustomObject]$Data
    }
    if ($AsJson) {
        return $Data | ConvertTo-Json -Depth 10 -Compress
    }
    return $Data
}

function Resolve-EnumValue {
    <#
    .SYNOPSIS
        Look up a friendly name in a PPT_* hashtable, returning the numeric enum value.
        If the key is already numeric, pass it through. Otherwise throw.
    #>
    param(
        [Parameter(Mandatory)]
        [hashtable]$Map,

        [Parameter(Mandatory)]
        [string]$Key
    )

    # Case-insensitive lookup
    foreach ($entry in $Map.GetEnumerator()) {
        if ($entry.Key -ieq $Key) {
            return $entry.Value
        }
    }

    # If key is numeric, pass through as-is
    if ($Key -match '^\s*-?\d+\s*$') {
        return [int]$Key
    }

    $validKeys = ($Map.Keys | Sort-Object) -join ', '
    throw "Unknown enum key '$Key'. Valid values: $validKeys"
}
