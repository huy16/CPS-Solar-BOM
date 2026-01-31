import XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

const filePath = path.resolve('../1. Database/TOOL_160925_BOQ_BOM_DNO(3 Rail).xlsm');

console.log(`Reading file: ${filePath}`);

try {
    // Read only the workbook structure first (not full data if possible, but xlsx reads file)
    // For 27MB, it mimics reading into memory. 
    const workbook = XLSX.readFile(filePath);

    const targetKeywords = ['data', 'equip', 'dashboard', 'boq tong'];

    console.log("Actual Keys in Workbook.Sheets:", Object.keys(workbook.Sheets));

    workbook.SheetNames.forEach(rawSheetName => {
        const sheetName = rawSheetName.trim();
        const lowerName = sheetName.toLowerCase();

        // Check if this sheet matches any keyword (fuzzy match)
        const isTarget = targetKeywords.some(k => lowerName.includes(k));

        if (isTarget) {
            console.log(`[Sheet Found in Loop: '${rawSheetName}']`);
            if (workbook.Sheets[rawSheetName]) {
                console.log("   -> Exists in Sheets object");
            } else {
                console.log("   -> DOES NOT EXIST in Sheets object");
            }
        }
    });

} catch (error) {
    console.error("Error reading file:", error.message);
}
