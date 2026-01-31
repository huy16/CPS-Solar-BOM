import ExcelJS from 'exceljs';

async function updateUnits() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('public/BOQ_BOM_Template.xlsx');

    const ws = wb.getWorksheet('BOM');

    // Mapping Vietnamese to English
    const unitMap = {
        'bộ': 'Set',
        'Bộ': 'Set',
        'bo': 'Set',
        'Mét': 'Meter',
        'mét': 'Meter',
        'm': 'Meter',
        'thùng': 'Box',
        'Thùng': 'Box',
        'cái': 'Pcs',
        'Cái': 'Pcs',
        'cây': 'Pcs',
        'Cây': 'Pcs',
        'cay': 'Pcs',
        'gói': 'Pack',
        'Gói': 'Pack',
        'cuộn': 'Roll',
        'Cuộn': 'Roll',
        'cuon': 'Roll',
        'ống': 'Tube',
        'Ống': 'Tube',
        'chai': 'Bottle',
        'sợi': 'Pcs',
        'lon': 'Can',
        'thanh': 'Pcs',
        'vị trí': 'Location'
    };

    let changes = [];

    // Column 5 is Unit column (E) - not column 4!
    for (let r = 14; r <= 130; r++) {
        const cell = ws.getRow(r).getCell(5);
        const val = cell.value;
        if (val && typeof val === 'string') {
            const trimmed = val.trim();
            if (unitMap[trimmed]) {
                changes.push({ row: r, old: trimmed, newVal: unitMap[trimmed] });
                cell.value = unitMap[trimmed];
            }
        }
    }

    console.log('Changes made:', changes.length);
    changes.forEach(c => console.log('Row ' + c.row + ': "' + c.old + '" -> "' + c.newVal + '"'));

    await wb.xlsx.writeFile('public/BOQ_BOM_Template.xlsx');
    console.log('Template updated successfully!');
}

updateUnits().catch(err => console.error('Error:', err));
