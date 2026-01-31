import ExcelJS from 'exceljs';

async function inspectHeaders() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('public/BOQ_BOM_Template.xlsx');

    const boq = workbook.getWorksheet('BOQ');
    const bom = workbook.getWorksheet('BOM');

    console.log("--- BOQ Headers (Row 3) ---");
    const boqRow = boq.getRow(3);
    for (let i = 1; i <= 25; i++) {
        console.log(`Col ${i}: ${boqRow.getCell(i).value}`);
    }

    console.log("\n--- BOM Headers (Row 14) ---");
    const bomRow = bom.getRow(14);
    for (let i = 1; i <= 13; i++) {
        console.log(`Col ${i}: ${bomRow.getCell(i).value}`);
    }
}

inspectHeaders();
