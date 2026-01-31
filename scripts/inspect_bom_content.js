import ExcelJS from 'exceljs';

async function inspectContent() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('public/BOQ_BOM_Template.xlsx');
    const ws = workbook.getWorksheet('BOM');

    console.log("--- BOM Content Rows 31-100 ---");
    for (let r = 31; r <= 100; r++) {
        const row = ws.getRow(r);
        const col1 = row.getCell(1).value;
        const col2 = row.getCell(2).value;
        const col3 = row.getCell(3).value;

        // Print only if meaningful (Header or Item)
        if (col1 || col2 || col3) {
            // Simplify output
            const c1 = col1 ? col1.toString().substring(0, 10) : "";
            const c2 = col2 ? col2.toString().replace(/\n/g, " ").substring(0, 50) : "";
            const c3 = col3 ? col3.toString().substring(0, 20) : "";
            console.log(`[${r}] ${c1} | ${c2} | ${c3}`);
        }
    }
}
inspectContent();
