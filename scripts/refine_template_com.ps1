$sourcePath = "D:\TOOL GOOGLE ANTIGRAVITY\6. Tool BOQ & BOM\1. Database\Template BOM&BOQ.xlsx"
$destPath = "D:\TOOL GOOGLE ANTIGRAVITY\6. Tool BOQ & BOM\boq-bom-app\public\BOQ_BOM_Template.xlsx"

Write-Host "Starting Excel COM..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Opening source template..."
    $workbook = $excel.Workbooks.Open($sourcePath)

    # 1. Clear Data in BOM Sheet
    try {
        $sheet = $workbook.Sheets.Item("BOM")
        Write-Host "Sheet 'BOM' found. Clearing data..."
        
        # Determine last row
        $lastRow = $sheet.Cells($sheet.Rows.Count, "A").End(-4162).Row # xlUp
        Write-Host "Last Row: $lastRow"

        if ($lastRow -ge 15) {
            # Clear Content from F15:M{lastRow}
            $rangeToClear = $sheet.Range("F15:M$lastRow")
            $rangeToClear.ClearContents()
            Write-Host "Cleared range F15:M$lastRow"
        }

        # Clear specific input that might be sample data?
        # If user wants "previous version", safe to assume they want the Template's original text.
        # We will stop touching the headers/labels entirely.



        # --- Footer Updates ---
        # Hardcoding rows based on static template inspection
        # Row 123, Col 2 (B): Location
        
        # Row 126, Col 9 (I): Date
        $sheet.Cells.Item(126, 9).Value2 = $strNgay
        
        Write-Host "Reset Footer: Location (Row 123) and Date (Row 126)"
    }
    catch {
        Write-Error "Sheet 'BOM' not found or error clearing: $_"
    }

    # 2. Clear Data in BOQ Sheet
    try {
        $sheetBOQ = $workbook.Sheets.Item("BOQ")
        Write-Host "Sheet 'BOQ' found. Clearing data from Row 4..."

        # Merge Row 1 (Title)
        $sheetBOQ.Range("A1:Y1").Merge()
        # $sheetBOQ.Range("A1:Y1").HorizontalAlignment = -4108 # xlCenter (Optional, keeps center)
        Write-Host "Merged A1:Y1 in BOQ"
        
        # Use Column B (STT) to find last row as Column A might be empty/merged
        $lastRowBOQ = $sheetBOQ.Cells($sheetBOQ.Rows.Count, "B").End(-4162).Row # xlUp
        Write-Host "BOQ Last Row (based on Col B): $lastRowBOQ"
        
        if ($lastRowBOQ -ge 4) {
            # Delete rows 4 to end to be clean
            $sheetBOQ.Rows.Item("4:$lastRowBOQ").Delete()
            Write-Host "Deleted rows 4 to $lastRowBOQ in BOQ"
        }
    }
    catch {
        Write-Error "Sheet 'BOQ' not found or error clearing: $_"
    }

    # 2. Save As Destination
    Write-Host "Saving to $destPath..."
    $workbook.SaveAs($destPath)
    Write-Host "Saved successfully."

}
catch {
    Write-Error "An error occurred: $_"
}
finally {
    Write-Host "Closing Excel..."
    if ($workbook) { $workbook.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel
}
