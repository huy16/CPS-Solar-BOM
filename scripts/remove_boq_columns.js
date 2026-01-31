import ExcelJS from 'exceljs';

async function removeColumns() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('public/BOQ_BOM_Template.xlsx');

    const ws = wb.getWorksheet('BOQ');

    console.log('Before - Headers row 3:');
    for (let c = 1; c <= 10; c++) {
        console.log('Col ' + c + ':', ws.getRow(3).getCell(c).value);
    }

    // Delete columns 1 and 2 (DTC and STT)
    // Need to delete from right to left to maintain correct indices
    ws.spliceColumns(2, 1); // Delete column 2 (STT)
    ws.spliceColumns(1, 1); // Delete column 1 (DTC)

    console.log('\nAfter - Headers row 3:');
    for (let c = 1; c <= 10; c++) {
        console.log('Col ' + c + ':', ws.getRow(3).getCell(c).value);
    }

    await wb.xlsx.writeFile('public/BOQ_BOM_Template.xlsx');
    console.log('\nTemplate updated - columns DTC and STT removed!');
}

removeColumns().catch(err => console.error('Error:', err));
