import * as XLSX from 'xlsx';
import fs from 'fs';

const filePath = "d:/TOOL GOOGLE ANTIGRAVITY/6. Tool BOQ & BOM/1. Database/Template BOM&BOQ.xlsx";

try {
    console.log(`Reading file: ${filePath}`);
    const buffer = fs.readFileSync(filePath);
    const wb = XLSX.read(buffer, { type: 'buffer' });

    wb.SheetNames.forEach(sheetName => {
        console.log(`\n--- Sheet: ${sheetName} ---`);
        const ws = wb.Sheets[sheetName];

        // Read first 20 rows
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, range: 0 });

        console.log(`Sheet Data Length: ${data.length}`);
        for (let i = 0; i < Math.min(20, data.length); i++) {
            // Only print if row is not empty
            if (data[i] && data[i].length > 0) {
                console.log(`Row ${i + 1}:`, JSON.stringify(data[i]));
            }
        }
    });

} catch (err) {
    console.error("Error reading file:", err);
}
