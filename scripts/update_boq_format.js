import ExcelJS from 'exceljs';

async function updateBOQFormat() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('public/BOQ_BOM_Template.xlsx');

    const ws = wb.getWorksheet('BOQ');

    // Increase column widths for J, Q, S
    ws.getColumn('J').width = 15;  // Tiếp địa Rail
    ws.getColumn('Q').width = 15;  // TIẾP ĐỊA C10
    ws.getColumn('S').width = 12;  // CÁP MẠNG

    console.log('Column widths updated:');
    console.log('J:', ws.getColumn('J').width);
    console.log('Q:', ws.getColumn('Q').width);
    console.log('S:', ws.getColumn('S').width);

    await wb.xlsx.writeFile('public/BOQ_BOM_Template.xlsx');
    console.log('Template updated!');
}

updateBOQFormat().catch(err => console.error('Error:', err));
