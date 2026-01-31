import ExcelJS from 'exceljs';
import path from 'path';

async function refineTemplate() {
    const inputPath = "d:/TOOL GOOGLE ANTIGRAVITY/6. Tool BOQ & BOM/1. Database/Template BOM&BOQ.xlsx";
    const outputPath = "public/BOQ_BOM_Template.xlsx";

    const workbook = new ExcelJS.Workbook();
    try {
        console.log(`Reading file: ${inputPath}`);
        await workbook.xlsx.readFile(inputPath);

        const sheetName = "BOM";
        const worksheet = workbook.getWorksheet(sheetName);

        if (!worksheet) {
            console.error(`Sheet '${sheetName}' not found!`);
            return;
        }

        console.log(`Processing sheet: ${sheetName}`);

        // Data starts from row 15 (based on previous inspection)
        // Headers are at row 13. 
        // Iterate all rows from 15 downwards
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber >= 15) {
                // Keep A-F (Indices 1-6 in ExcelJS 1-based)
                // Clear G onwards (Index 7+)
                // Columns: A(1), B(2), C(3), D(4), E(5), F(6) -> KEEP
                // G(7), H(8), I(9), J(10), K(11), L(12), M(13) -> CLEAR

                for (let col = 7; col <= 20; col++) { // Clear up to col 20 just in case
                    const cell = row.getCell(col);
                    cell.value = null;
                }
            }
        });

        console.log(`Writing refined template to: ${outputPath}`);
        await workbook.xlsx.writeFile(outputPath);
        console.log("Done.");

    } catch (err) {
        console.error("Error:", err);
    }
}

refineTemplate();
