
import ExcelJS from 'exceljs';
import path from 'path';

async function checkDetails() {
    const workbook = new ExcelJS.Workbook();
    const filePath = path.resolve('public/BOQ_BOM_Template.xlsx');
    await workbook.xlsx.readFile(filePath);
    const ws = workbook.getWorksheet("BOM");

    const romanRegex = /^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII)$/;

    let currentGroup = null;
    const groups = [];

    // Safe text extractor
    const getText = (v) => {
        if (!v) return "";
        if (typeof v === 'string' || typeof v === 'number') return v.toString();

        // Handle Rich Text
        if (v.richText && Array.isArray(v.richText)) {
            return v.richText.map(t => t.text).join("");
        }

        // Handle Hyperlink or other objects
        if (v.text) return v.text;

        // Fallback for unknown objects
        try {
            return JSON.stringify(v);
        } catch (e) {
            return "[Complex Object]";
        }
    };

    ws.eachRow((row, rowNumber) => {
        const col1 = row.getCell(1).value;
        const col2 = row.getCell(2).value;
        const col3 = row.getCell(3).value;

        // Check for Group Header
        if (typeof col1 === 'string' && romanRegex.test(col1.trim())) {
            currentGroup = {
                id: col1.trim(),
                name: getText(col2),
                items: []
            };
            groups.push(currentGroup);
        } else if (currentGroup) {
            // Check for footer to stop
            const str2 = getText(col2);
            if (str2.includes("Hồ sơ") || str2.includes("Địa điểm") || str2.includes("Ngày")) {
                currentGroup = null;
                return;
            }

            // Capture Item
            const nameVal = getText(col2).trim();
            const codeVal = getText(col3).trim();

            if (nameVal || codeVal) {
                currentGroup.items.push({
                    name: nameVal,
                    code: codeVal
                });
            }
        }
    });

    // Print Report
    groups.forEach(g => {
        console.log(`\n### ${g.id} - ${g.name}`);
        g.items.forEach(item => {
            const codeStr = item.code ? ` [${item.code}]` : "";
            // Clean newline characters which might mess up console
            const cleanName = item.name.replace(/\n/g, " ");
            console.log(`- ${cleanName}${codeStr}`);
        });
    });
}

checkDetails();
