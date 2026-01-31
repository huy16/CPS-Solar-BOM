$excelPath = "D:\TOOL GOOGLE ANTIGRAVITY\6. Tool BOQ & BOM\1. Database\TOOL_160925_BOQ_BOM_DNO(3 Rail).xlsm"
$jsonPath = "D:\TOOL GOOGLE ANTIGRAVITY\6. Tool BOQ & BOM\boq-bom-app\src\data\equipment_data.json"

Write-Host "Starting Excel COM..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Opening workbook..."
    $workbook = $excel.Workbooks.Open($excelPath)
    
    try {
        $sheet = $workbook.Sheets.Item("DATA EQUIP")
        Write-Host "Sheet 'DATA EQUIP' found."
    } catch {
        Write-Error "Sheet 'DATA EQUIP' not found!"
        exit 1
    }

    $lastRow = $sheet.Cells($sheet.Rows.Count, "A").End(-4162).Row # xlUp = -4162
    Write-Host "Last Row: $lastRow"

    # Helper function to find column index
    function Get-ColumnIndex {
        param([string]$name)
        try {
            # Search in first row
            $found = $sheet.Rows.Item(1).Find($name)
            if ($found -ne $null) {
                return $found.Column
            }
        } catch {}
        return 0
    }

    $invModelCol = Get-ColumnIndex "Inv Model"
    $invMinVolCol = Get-ColumnIndex "Min Voltage Operating"
    $invMaxVolCol = Get-ColumnIndex "Max voltage Operating"
    $invNumInputsCol = Get-ColumnIndex "Num Inputs"
    $invCapacityCol = Get-ColumnIndex "Capacity"

    $pvModelCol = Get-ColumnIndex "PV Model"
    # Mapping AM -> 39, AP -> 42 (1-based)
    $pvCapacityCol = 39 
    $pvVocCol = 42

    Write-Host "Columns: InvModel=$invModelCol, PVModel=$pvModelCol"

    $inverters = @{}
    $pvs = @{}

    # Read Data using range for speed (instead of cell by cell)
    # But cell by cell is safer/easier for sparse logic in PS. For <1000 rows it's fine.
    
    for ($r = 2; $r -le $lastRow; $r++) {
        if ($r % 50 -eq 0) { Write-Host "Processing row $r..." }

        # Inverter
        if ($invModelCol -gt 0) {
            $invName = [string]$sheet.Cells.Item($r, $invModelCol).Text
            $invName = $invName.Trim()
            if ($invName -ne "") {
                if (-not $inverters.ContainsKey($invName)) {
                    $inverters[$invName] = @{
                        "minVoltage" = $sheet.Cells.Item($r, $invMinVolCol).Value2
                        "maxVoltage" = $sheet.Cells.Item($r, $invMaxVolCol).Value2
                        "numInputs"  = $sheet.Cells.Item($r, $invNumInputsCol).Value2
                        "capacityKw" = $sheet.Cells.Item($r, $invCapacityCol).Value2
                    }
                }
            }
        }

        # PV
        if ($pvModelCol -gt 0) {
            $pvName = [string]$sheet.Cells.Item($r, $pvModelCol).Text
            $pvName = $pvName.Trim()
            if ($pvName -ne "") {
                if (-not $pvs.ContainsKey($pvName)) {
                    $cap = $sheet.Cells.Item($r, $pvCapacityCol).Value2
                    $voc = $sheet.Cells.Item($r, $pvVocCol).Value2
                    
                    if ($cap -is [double] -or $cap -is [int]) { $cap = $cap / 1000 }
                    
                    $pvs[$pvName] = @{
                        "powerKwp" = $cap
                        "voc"      = $voc
                    }
                }
            }
        }
    }

    $output = @{
        "inverters" = $inverters
        "photovoltaics" = $pvs
    }

    Write-Host "Saving JSON..."
    $json = $output | ConvertTo-Json -Depth 5
    Set-Content -Path $jsonPath -Value $json

    Write-Host "Done!"

} catch {
    Write-Error "An error occurred: $_"
} finally {
    Write-Host "Closing Excel..."
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel
}
