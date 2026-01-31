import ExcelJS from 'exceljs';

async function scanGroups() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('public/BOQ_BOM_Template.xlsx');
    const ws = workbook.getWorksheet('BOM');

    const groups = {};
    ws.eachRow((row, rowNumber) => {
        const val = row.getCell(1).value; // Column A
        // Check for Roman Numerals I through XIII
        if (typeof val === 'string' && /^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII)$/.test(val)) {
            console.log(`Found Group ${val} at Row ${rowNumber} - Label: ${row.getCell(2).value}`);
        }
    });
}
scanGroups();
