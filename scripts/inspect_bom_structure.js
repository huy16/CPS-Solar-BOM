import ExcelJS from 'exceljs';

async function inspectStructure() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('public/BOQ_BOM_Template.xlsx');
    const ws = workbook.getWorksheet('BOM');

    console.log("--- BOM Structure Rows 10-30 ---");
    for (let r = 10; r <= 30; r++) {
        const row = ws.getRow(r);
        const values = [];
        for (let c = 1; c <= 6; c++) { // Top 6 cols
            values.push(row.getCell(c).value);
        }
        console.log(`Row ${r}: ${JSON.stringify(values)}`);
    }
}

inspectStructure();
