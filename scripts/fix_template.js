import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

async function fixTemplate() {
    const filePath = path.join(__dirname, '../public/BOQ_BOM_Template.xlsx');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const sheetBOM = workbook.getWorksheet("BOM");
    if (!sheetBOM) {
        console.error("Sheet BOM not found!");
        return;
    }

    console.log("Fixing BOM labels...");
    // C6: " - Mã tủ:"
    sheetBOM.getCell('C6').value = " - Mã tủ:";
    // C7: " - Mã dự án:"
    sheetBOM.getCell('C7').value = " - Mã dự án:";
    // C8: " - Hạng mục:"
    sheetBOM.getCell('C8').value = " - Hạng mục:";
    // C9: " - Lý do mua hàng:"
    sheetBOM.getCell('C9').value = " - Lý do mua hàng:";

    // I9: " - Mã dự án:"
    sheetBOM.getCell('I9').value = " - Mã dự án:";
    // I10: " - Hạng mục:"
    sheetBOM.getCell('I10').value = " - Hạng mục:";

    // Footer
    // B123
    sheetBOM.getCell('B123').value = "Địa điểm yêu cầu:";
    // I126
    sheetBOM.getCell('I126').value = "Ngày: .../.../...";

    // Set Fonts to Times New Roman just to be safe
    ['C6', 'C7', 'C8', 'C9', 'I9', 'I10', 'B123', 'I126'].forEach(cell => {
        sheetBOM.getCell(cell).font = { name: 'Times New Roman', size: 11, bold: false }; // Match template style roughly
    });


    await workbook.xlsx.writeFile(filePath);
    console.log("Template fixed successfully!");
}

fixTemplate();
