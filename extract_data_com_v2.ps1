$excelPath = "D:\TOOL GOOGLE ANTIGRAVITY\6. Tool BOQ & BOM\1. Database\TOOL_160925_BOQ_BOM_DNO(3 Rail).xlsm"
$jsonPath = "D:\TOOL GOOGLE ANTIGRAVITY\6. Tool BOQ & BOM\boq-bom-app\src\data\equipment_data.json"

Write-Host "Starting Excel COM (V2)..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbook = $excel.Workbooks.Open($excelPath)
    $sheet = $workbook.Sheets.Item("DATA EQUIP")
    $lastRow = $sheet.Cells($sheet.Rows.Count, "A").End(-4162).Row # xlUp
    Write-Host "Last Row: $lastRow"

    if ($lastRow -lt 2) { 
        Write-Error "Not enough data"
        exit 1
    }

    # Find Columns
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

    # Read All Data
    # Determine max column needed (45 is enough for AP=42)
    # InvModel is 66 (BN). We need more.
    $maxCol = 100 
    $lastColLetter = "CV" # roughly
    $range = $sheet.Range("A1", $sheet.Cells.Item($lastRow, $maxCol))
    Write-Host "Reading Range..."
    $data = $range.Value2
    
    Write-Host "Data Type: $($data.GetType().Name)"
    Write-Host "Rank: $($data.Rank)"
    
    $inverters = @{}
    $pvs = @{}

    # Loop rows
    # 2D Array bounds in COM are usually [1..Rows, 1..Cols]
    $rowCount = $data.GetUpperBound(0)
    
    for ($r = 2; $r -le $rowCount; $r++) {
        # Inverter
        if ($invModelIdx -gt 0) {
            $m = $data[$r, $invModelIdx]
            if ($m -ne $null) {
                $m = $m.ToString().Trim()
                if ($m -ne "" -and -not $inverters.ContainsKey($m)) {
                    $inverters[$m] = @{
                        "minVoltage" = $data[$r, $invMinVolIdx]
                        "maxVoltage" = $data[$r, $invMaxVolIdx]
                        "numInputs"  = $data[$r, $invNumInputsIdx]
                        "capacityKw" = $data[$r, $invCapacityIdx]
                    }
                }
            }
        }

        # PV
        if ($pvModelIdx -gt 0) {
            $p = $data[$r, $pvModelIdx]
            if ($p -ne $null) {
                $p = $p.ToString().Trim()
                if ($p -ne "" -and -not $pvs.ContainsKey($p)) {
                    $cap = $data[$r, $pvCapacityIdx]
                    $voc = $data[$r, $pvVocIdx]

                    # Convert cap to double safely
                    try { $cap = [double]$cap } catch { $cap = 0 }
                    
                    if ($cap -gt 10) { $cap = $cap / 1000 }

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
    Write-Host "Inverters: $($inverters.Count)"
    Write-Host "PVs: $($pvs.Count)"
    
    # Peek a key to verify
    if ($pvs.Count -gt 0) {
        Write-Host "Sample Key: $($pvs.Keys[0])"
    }

}
catch {
    Write-Error "Error: $_"
}
finally {
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
