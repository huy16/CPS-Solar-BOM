import ExcelJS from 'exceljs';

async function inspect() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('public/BOQ_BOM_Template.xlsx');
    const ws = workbook.getWorksheet('BOM');

    // Row 15 is the data row placeholder
    const row = ws.getRow(15);
    const cell = row.getCell(2); // Column B (2)

    console.log("--- Style Inspection Row 15, Col 2 ---");
    console.log("Font:", cell.font);
    console.log("Border:", cell.border);
    console.log("Alignment:", cell.alignment);
}

inspect();
