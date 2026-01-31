import XLSX from 'xlsx';
import path from 'path';

// Target a Datalink file
const filePath = path.resolve('../1. Database/Project/Datalink/1. DATALINK_DNA.xlsx');

console.log(`Reading Datalink file: ${filePath}`);

try {
    const workbook = XLSX.readFile(filePath);

    console.log("--- Sheet Names ---");
    console.log(workbook.SheetNames);

    // Usually Datalink files have a main sheet. Let's look at the first one and any "Data" one.
    workbook.SheetNames.slice(0, 3).forEach(sheetName => {
        console.log(`\n[Sheet: '${sheetName}']`);
        const sheet = workbook.Sheets[sheetName];

        // Preview top 20 rows, 20 cols
        const preview = XLSX.utils.sheet_to_json(sheet, { header: 1, range: "A1:T20" });
        console.log(JSON.stringify(preview, null, 2));
    });

} catch (error) {
    console.error("Error reading file:", error.message);
}
