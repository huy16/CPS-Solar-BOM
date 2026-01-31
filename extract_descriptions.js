
import ExcelJS from 'exceljs';
import path from 'path';

async function extract() {
    const workbook = new ExcelJS.Workbook();
    // Force reading the public template
    const filePath = path.resolve('public/BOQ_BOM_Template.xlsx');
    await workbook.xlsx.readFile(filePath);
    const ws = workbook.getWorksheet("BOM");

    // Groups I want: II (Panel), III (Inverter)
    const targetGroups = ['II', 'III'];
    let currentGid = null;

    const results = {};

    // Helper
    const getText = (v) => {
        if (!v) return "";
        if (typeof v === 'string' || typeof v === 'number') return v.toString();
        if (v.richText && Array.isArray(v.richText)) {
            return v.richText.map(t => t.text).join("");
        }
        if (v.text) return v.text;
        return "";
    };

    ws.eachRow((row, rowNumber) => {
        const col1 = row.getCell(1).value;
        const col2 = row.getCell(2).value; // Desc
        const col3 = row.getCell(3).value; // Code
        const col4 = row.getCell(4).value; // Supplier
        const col5 = row.getCell(5).value; // Unit

        // Detect Group Header
        if (typeof col1 === 'string') {
            const val = col1.trim();
            if (/^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII)$/.test(val)) {
                currentGid = val;
                return;
            }
        }

        // Capture Data if in target group
        if (currentGid && targetGroups.includes(currentGid)) {
            // Stop at footer
            const desc = getText(col2);
            if (desc.includes("Hồ sơ") || desc.includes("Địa điểm")) {
                currentGid = null;
                return;
            }

            const code = getText(col3).trim();
            if (code) {
                results[code] = {
                    description: desc.trim(),
                    supplier: getText(col4).trim(),
                    unit: getText(col5).trim()
                };
            }
        }
    });

    console.log(JSON.stringify(results, null, 2));
}

extract();
