import * as XLSX from 'xlsx';
import fs from 'fs';

const filePath = "d:/TOOL GOOGLE ANTIGRAVITY/6. Tool BOQ & BOM/1. Database/Template BOM&BOQ.xlsx";

try {
    console.log(`Reading file: ${filePath}`);
    const buffer = fs.readFileSync(filePath);
    const wb = XLSX.read(buffer, { type: 'buffer' });

    const sheetName = "BOM";
    const ws = wb.Sheets[sheetName];

    if (!ws) {
        console.log("Sheet BOM not found");
        process.exit(1);
    }

    const range = XLSX.utils.decode_range(ws['!ref']);
    const totalRows = range.e.r + 1;
    console.log(`Total Rows: ${totalRows}`);

    const startRow = Math.max(0, totalRows - 20);
    console.log(`Reading from row ${startRow + 1} to ${totalRows}`);

    // Read full sheet to json but we slice manualy or use options
    // Using sheet_to_json with range option is safer

    const data = XLSX.utils.sheet_to_json(ws, {
        header: 1,
        range: { s: { r: startRow, c: 0 }, e: { r: totalRows, c: 20 } }
    });

    data.forEach((row, idx) => {
        if (row && row.length > 0) {
            console.log(`Row ${startRow + idx + 1}:`, JSON.stringify(row));
        }
    });

} catch (err) {
    console.error("Error reading file:", err);
}
