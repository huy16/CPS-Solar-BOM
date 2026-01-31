
import ExcelJS from 'exceljs';
import path from 'path';

async function checkGroups() {
    const workbook = new ExcelJS.Workbook();
    const filePath = path.resolve('public/BOQ_BOM_Template.xlsx');
    await workbook.xlsx.readFile(filePath);
    const ws = workbook.getWorksheet("BOM");

    const romanRegex = /^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII)$/;

    console.log("Groups Found:");
    ws.eachRow((row, rowNumber) => {
        const val = row.getCell(1).value;
        if (typeof val === 'string' && romanRegex.test(val.trim())) {
            const name = row.getCell(2).value;
            console.log(`${val.trim()}: ${name}`);
        }
    });
}

checkGroups();
