$excelPath = "D:\TOOL GOOGLE ANTIGRAVITY\6. Tool BOQ & BOM\1. Database\TOOL_160925_BOQ_BOM_DNO(3 Rail).xlsm"
$jsonPath = "D:\TOOL GOOGLE ANTIGRAVITY\6. Tool BOQ & BOM\boq-bom-app\src\data\equipment_data.json"

Write-Host "Starting Excel COM (Fast)..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($excelPath)
    $sheet = $workbook.Sheets.Item("DATA EQUIP")
    $lastRow = $sheet.Cells($sheet.Rows.Count, "A").End(-4162).Row
    Write-Host "Rows to process: $lastRow"

    if ($lastRow -lt 2) { $lastRow = 2 }

    # Function to get column letter from number
    function Get-ColLetter($intCol) {
        if ($intCol -le 0) { return "" }
        $v = $intCol
        $str = ""
        while ($v -gt 0) {
            $m = ($v - 1) % 26
            $str = [char][int](65 + $m) + $str
            $v = [math]::Floor(($v - $m) / 26)
        }
        return $str
    }

    function Get-ColumnIndex {
        param([string]$name)
        try {
            $found = $sheet.Rows.Item(1).Find($name)
            if ($found -ne $null) { return $found.Column }
        }
        catch {}
        return 0
    }

    $invModelIdx = Get-ColumnIndex "Inv Model"
    $invMinVolIdx = Get-ColumnIndex "Min Voltage Operating"
    $invMaxVolIdx = Get-ColumnIndex "Max voltage Operating"
    $invNumInputsIdx = Get-ColumnIndex "Num Inputs"
    $invCapacityIdx = Get-ColumnIndex "Capacity"

    $pvModelIdx = Get-ColumnIndex "PV Model"
    $pvCapacityIdx = 39 # AM
    $pvVocIdx = 42      # AP

    Write-Host "Indices: InvModel=$invModelIdx, PVModel=$pvModelIdx"

    # Read columns into arrays (2D arrays returned by Value2)
    function Read-Column($colIdx) {
        if ($colIdx -le 0) { return $null }
        $letter = Get-ColLetter $colIdx
        $rangeStr = "$letter`2:$letter$lastRow"
        Write-Host "Reading $rangeStr ..."
        try {
            return $sheet.Range($rangeStr).Value2
        }
        catch {
            Write-Warning "Failed to read column $colIdx"
            return $null
        }
    }

    $invModels = Read-Column $invModelIdx
    $invMinVols = Read-Column $invMinVolIdx
    $invMaxVols = Read-Column $invMaxVolIdx
    $invNumInputs = Read-Column $invNumInputsIdx
    $invCaps = Read-Column $invCapacityIdx

    $pvModels = Read-Column $pvModelIdx
    $pvCaps = Read-Column $pvCapacityIdx
    $pvVocs = Read-Column $pvVocIdx

    $inverters = @{}
    $pvs = @{}

    # Process arrays
    # Value2 returns 2D array [row, col] (1-based), but in PS it might be treated as multidimensional array
    # Accessing: $invModels[row, 1] 
    
    Write-Host "Processing data in memory..."
    
    for ($i = 1; $i -le ($lastRow - 1); $i++) {
        # Inverter
        if ($invModels -ne $null) {
            $m = $invModels[$i, 1]
            if ($m -ne $null -and $m -ne "") {
                $m = $m.ToString().Trim()
                if (-not $inverters.ContainsKey($m)) {
                    $inverters[$m] = @{
                        "minVoltage" = if ($invMinVols) { $invMinVols[$i, 1] } else { 0 }
                        "maxVoltage" = if ($invMaxVols) { $invMaxVols[$i, 1] } else { 0 }
                        "numInputs"  = if ($invNumInputs) { $invNumInputs[$i, 1] } else { 0 }
                        "capacityKw" = if ($invCaps) { $invCaps[$i, 1] } else { 0 }
                    }
                }
            }
        }

        # PV
        if ($pvModels -ne $null) {
            $p = $pvModels[$i, 1]
            if ($p -ne $null -and $p -ne "") {
                $p = $p.ToString().Trim()
                if (-not $pvs.ContainsKey($p)) {
                    $cap = if ($pvCaps) { $pvCaps[$i, 1] } else { 0 }
                    $voc = if ($pvVocs) { $pvVocs[$i, 1] } else { 0 }
                    
                    # Ensure scalar
                    if ($cap -is [Array]) { $cap = $cap[0] }
                    
                    # Convert to double
                    try {
                        $cap = [double]$cap
                    }
                    catch {
                        $cap = 0
                    }

                    if ($cap -gt 10) { $cap = $cap / 1000 } # Assume >10 is Watts, convert to kW

                    $pvs[$p] = @{
                        "powerKwp" = $cap
                        "voc"      = $voc
                    }
                }
            }
        }
    }

    $output = @{
        "inverters"     = $inverters
        "photovoltaics" = $pvs
    }

    $json = $output | ConvertTo-Json -Depth 5
    Set-Content -Path $jsonPath -Value $json
    Write-Host "Saved JSON to $jsonPath"
    Write-Host "Inverters found: $($inverters.Count)"
    Write-Host "PV Models found: $($pvs.Count)"

}
catch {
    Write-Error "Error: $_"
}
finally {
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
