import XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

// Define paths
// Using the path from the user information
const EXCEL_PATH = 'D:\\TOOL GOOGLE ANTIGRAVITY\\6. Tool BOQ & BOM\\1. Database\\TOOL_160925_BOQ_BOM_DNO(3 Rail).xlsm';
const OUTPUT_PATH = 'D:\\TOOL GOOGLE ANTIGRAVITY\\6. Tool BOQ & BOM\\boq-bom-app\\src\\data\\equipment_data.json';

function extractEquipmentData() {
    console.log(`Reading Excel file from: ${EXCEL_PATH}`);

    if (!fs.existsSync(EXCEL_PATH)) {
        console.error(`Error: File not found at ${EXCEL_PATH}`);
        process.exit(1);
    }

    const workbook = XLSX.readFile(EXCEL_PATH);
    // Robust finding of sheet name using includes
    const sheetName = workbook.SheetNames.find(n => n.toUpperCase().includes('DATA') && n.toUpperCase().includes('EQUIP'));

    if (!sheetName) {
        console.error("Error: Sheet containing 'DATA' and 'EQUIP' not found in workbook.SheetNames!");
        // Log all names for debugging
        console.log("Available SheetNames:", workbook.SheetNames);
        process.exit(1);
    }

    console.log(`Found Sheet Name: "${sheetName}"`);
    console.log(`Type: ${typeof sheetName}, Length: ${sheetName.length}`);
    for (let i = 0; i < sheetName.length; i++) {
        console.log(`Char at ${i}: ${sheetName.charCodeAt(i)}`);
    }

    let worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
        console.log("Direct access failed. Iterating keys...");
        const keys = Object.keys(workbook.Sheets);
        const matchedKey = keys.find(k => k === sheetName);
        console.log(`Key exact match found: ${matchedKey}`);

        if (!matchedKey) {
            console.log("Trying fuzzy match on keys...");
            const fuzzyKey = keys.find(k => k.includes('EQUIP'));
            console.log(`Fuzzy key found: "${fuzzyKey}"`);
            if (fuzzyKey) worksheet = workbook.Sheets[fuzzyKey];
        }
    }

    if (!worksheet) {
        console.error("Critical Error: Sheet name found in list but Worksheet object is undefined.");
        process.exit(1);
    }
    console.log("Sheet Range:", worksheet['!ref']);
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    console.log("First 5 rows:", JSON.stringify(data.slice(0, 5)));

    if (data.length < 2) {
        console.error("Error: No data found in 'DATA EQUIP' sheet.");
        process.exit(1);
    }

    const headers = data[0].map(h => (h ? h.toString().trim() : ''));
    console.log("Headers found:", headers);

    // Helper to find column index (fuzzy match)
    const getColIndex = (name) => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

    // Specific column indices based on VBA logic
    const invModelIdx = getColIndex("Inv Model");
    const invMinVolIdx = getColIndex("Min Voltage Operating");
    const invMaxVolIdx = getColIndex("Max voltage Operating");
    const invNumInputsIdx = getColIndex("Num Inputs");
    const invCapacityIdx = getColIndex("Capacity");

    const pvModelIdx = getColIndex("PV Model");
    // VBA says AM and AP for Capacity and Voc. AM is index 38, AP is index 41 (0-indexed: A=0, Z=25, AA=26... )
    // A..Z (26) + A..M (13) = 39. So AM is column 39 (1-based) -> index 38.
    // A..Z (26) + A..P (16) = 42. So AP is column 42 (1-based) -> index 41.
    const pvCapacityIdx = 38;
    const pvVocIdx = 41;

    console.log("Column Indices:", {
        invModelIdx, invMinVolIdx, invMaxVolIdx, invNumInputsIdx, invCapacityIdx,
        pvModelIdx, pvCapacityIdx, pvVocIdx
    });

    const equipmentData = {
        inverters: {},
        photovoltaics: {} // Renamed to photovoltaics for clarity, originally pvData
    };

    for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // Extract Inverter Data
        if (invModelIdx !== -1 && row[invModelIdx]) {
            const model = row[invModelIdx].toString().trim();
            if (model) {
                equipmentData.inverters[model] = {
                    minVoltage: row[invMinVolIdx],
                    maxVoltage: row[invMaxVolIdx],
                    numInputs: row[invNumInputsIdx],
                    capacityKw: row[invCapacityIdx]
                };
            }
        }

        // Extract PV Data
        if (pvModelIdx !== -1 && row[pvModelIdx]) {
            const model = row[pvModelIdx].toString().trim();
            if (model) {
                // VBA logic: Capacity is divided by 1000 to get kWp
                // Ensure values are treated as numbers
                const rawCapacity = row[pvCapacityIdx];
                const rawVoc = row[pvVocIdx];

                let capacity = 0;
                let voc = 0;

                if (typeof rawCapacity === 'number') capacity = rawCapacity / 1000;
                else if (rawCapacity) capacity = parseFloat(rawCapacity) / 1000;

                if (typeof rawVoc === 'number') voc = rawVoc;
                else if (rawVoc) voc = parseFloat(rawVoc);

                equipmentData.photovoltaics[model] = {
                    powerKwp: capacity || 0, // Fallback to 0 if NaN
                    voc: voc || 0
                };
            }
        }
    }

    // Ensure output directory exists
    const outputDir = path.dirname(OUTPUT_PATH);
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true });
    }

    fs.writeFileSync(OUTPUT_PATH, JSON.stringify(equipmentData, null, 2));
    console.log(`Successfully extracted equipment data to ${OUTPUT_PATH}`);
    console.log(`Inverters found: ${Object.keys(equipmentData.inverters).length}`);
    console.log(`PV Models found: ${Object.keys(equipmentData.photovoltaics).length}`);
}

extractEquipmentData();
