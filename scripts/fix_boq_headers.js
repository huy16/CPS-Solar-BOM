import ExcelJS from 'exceljs';

async function fixBOQHeaders() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('public/BOQ_BOM_Template.xlsx');

    const ws = wb.getWorksheet('BOQ');

    // First, unmerge all cells in rows 1-3
    console.log('Unmerging cells...');
    const mergesToRemove = [...ws.model.merges];
    mergesToRemove.forEach(m => {
        try {
            ws.unMergeCells(m);
        } catch (e) {
            console.log('Could not unmerge:', m);
        }
    });

    // Define new headers structure (after removing DTC and STT)
    // Row 2: Group headers
    // Row 3: Individual column headers

    const headerRow2 = [
        'MST',           // A
        'TÊN DỰ ÁN',     // B
        'VIMETCO STRUCTURES', // C-J (merged)
        '', '', '', '', '', '', '',
        'HUAWEI',        // K-O (merged)
        '', '', '', '',
        'CÁP',           // P-S (merged)
        '', '', '',
        'KHÁC',          // T-V (merged)
        '', '',
        'NOTE'           // W
    ];

    const headerRow3 = [
        'MST',           // A
        'TÊN DỰ ÁN',     // B
        'Rail',          // C
        'Nối Rail',      // D
        'Chân L',        // E
        'Kẹp cuối',      // F
        'Kẹp giữa',      // G
        'Kẹp dây',       // H
        'Tiếp địa PIN',  // I
        'Tiếp địa Rail', // J
        'PV',            // K
        'MC4',           // L
        'INVERTER',      // M
        'SDONGLE',       // N
        'METER',         // O
        'CÁP AC',        // P
        'TIẾP ĐỊA C10',  // Q
        'CÁP DC',        // R
        'CÁP MẠNG',      // S
        'Ruột gà',       // T
        'JUNC_BOX',      // U
        'SUNTREE',       // V
        'NOTE'           // W
    ];

    // Clear and set Row 1 (Title)
    ws.getRow(1).getCell(1).value = '2. BILL OF QUANTITIES FOR INSTALLATION';
    ws.mergeCells('A1:W1');
    ws.getRow(1).getCell(1).style = {
        font: { bold: true, size: 14 },
        alignment: { horizontal: 'center', vertical: 'middle' }
    };

    // Set Row 2
    for (let i = 0; i < headerRow2.length; i++) {
        ws.getRow(2).getCell(i + 1).value = headerRow2[i];
    }

    // Set Row 3
    for (let i = 0; i < headerRow3.length; i++) {
        ws.getRow(3).getCell(i + 1).value = headerRow3[i];
    }

    // Merge cells for group headers in row 2
    ws.mergeCells('A2:A3'); // MST
    ws.mergeCells('B2:B3'); // TÊN DỰ ÁN
    ws.mergeCells('C2:J2'); // CLENERGY STRUCTURES
    ws.mergeCells('K2:O2'); // HUAWEI
    ws.mergeCells('P2:S2'); // CÁP
    ws.mergeCells('T2:V2'); // KHÁC
    ws.mergeCells('W2:W3'); // NOTE

    // Apply header style
    const headerStyle = {
        font: { bold: true },
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
        border: {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        },
        fill: {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        }
    };

    for (let c = 1; c <= 23; c++) {
        ws.getRow(2).getCell(c).style = headerStyle;
        ws.getRow(3).getCell(c).style = headerStyle;
    }

    // Set row heights
    ws.getRow(2).height = 25;
    ws.getRow(3).height = 25;

    await wb.xlsx.writeFile('public/BOQ_BOM_Template.xlsx');
    console.log('BOQ headers fixed!');

    // Verify
    console.log('\nNew merged cells:');
    ws.model.merges.forEach(m => console.log(m));
}

fixBOQHeaders().catch(err => console.error('Error:', err));
