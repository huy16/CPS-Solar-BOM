
import ExcelJS from 'exceljs';
import path from 'path';

async function generateTemplate() {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Input Data');

    // Define Columns
    ws.columns = [
        { header: 'Mã Trạm (Station Code)', key: 'code', width: 25 },
        { header: 'Công Suất (Capacity kWp)', key: 'power', width: 25 },
        { header: 'Tên Dự Án (Project Name)', key: 'name', width: 30 },
        { header: 'Ghi Chú (Note)', key: 'note', width: 20 }
    ];

    // Style Header
    ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    ws.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4F81BD' }
    };

    // Add Sample Data
    ws.addRow({ code: 'DMX_EXAMPLE_01', power: 15.5, name: 'Shop DMX 01 Example', note: 'Mẫu nhập liệu' });
    ws.addRow({ code: 'BHX_EXAMPLE_02', power: 20.0, name: 'Shop BHX 02 Example', note: '' });

    // Save to public folder
    const outputPath = path.resolve('public/Input_Template.xlsx');
    await workbook.xlsx.writeFile(outputPath);
    console.log(`Template generated at: ${outputPath}`);
}

generateTemplate();
