import XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';

const filePath = path.resolve('../1. Database/TOOL_160925_BOQ_BOM_DNO(3 Rail).xlsm');
const outputPath = path.resolve('src/data/sources/materials.json');

console.log(`Reading materials from: ${filePath}`);

try {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets['BOQ Tong ']; // Note space

    if (!sheet) {
        throw new Error("Sheet 'BOQ Tong ' not found or empty");
    }

    // Row 3 (Index 2) is Category/Name
    // Row 4 (Index 3) is Part Number/Model
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "D3:U5" });

    // headerRow = ["Rail", "Noi Rail", "Chan L", ...]
    // modelRow = ["ER-R-ECO/5200", "ER-SP-ECO", ...]

    // Rows are 0-indexed in the 'data' array we just grabbed
    // D3 is row index 0 relative to range D3:U5? No, range returns the block.
    // Let's grab the whole sheet slightly generously
    const fullData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "A1:AA10" });

    const categories = fullData[2]; // Row 3
    const models = fullData[3];     // Row 4

    const materials = [];

    // Indices correspond to columns. Let's find where "Rail" starts. 
    // In fullData, Rail is at index 3 (Column D).

    for (let i = 3; i < categories.length; i++) {
        const name = categories[i];
        const model = models[i];

        if (name && model) {
            materials.push({
                id: `mat_${i}`,
                name: name,
                model: model,
                type: deriveType(name, model),
                unit: 'pcs', // Default, logic to refine later
                unitPrice: 0 // Placeholder
            });
        }
    }

    // Add Inverter/Panel placeholders inferred from other columns if needed
    // But the columns D->... seem to cover the main structure.

    console.log(`Found ${materials.length} materials.`);
    fs.writeFileSync(outputPath, JSON.stringify(materials, null, 2));
    console.log(`Saved to ${outputPath}`);

} catch (error) {
    console.error("Error:", error.message);
}

function deriveType(name, model) {
    const n = name.toLowerCase();
    const m = model.toLowerCase();
    if (n.includes('rail')) return 'Rail';
    if (n.includes('kep') || n.includes('clamp')) return 'Clamp';
    if (n.includes('chan') || n.includes('foot')) return 'L-Foot';
    if (n.includes('cap') || n.includes('cable')) return 'Cable';
    if (n.includes('tiep dia') || n.includes('ground')) return 'Grounding';
    if (n.includes('inverter')) return 'Inverter';
    return 'Pars'; // Default
}
