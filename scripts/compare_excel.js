import ExcelJS from 'exceljs';

async function compare() {
    console.log("Loading files...");
    const workbookExport = new ExcelJS.Workbook();
    await workbookExport.xlsx.readFile('D:/TOOL GOOGLE ANTIGRAVITY/6. Tool BOQ & BOM/1. Database/BOQ_Export_2026-01-31.xlsx');

    const workbookTemplate = new ExcelJS.Workbook();
    await workbookTemplate.xlsx.readFile('D:/TOOL GOOGLE ANTIGRAVITY/6. Tool BOQ & BOM/1. Database/Template BOM&BOQ.xlsx');

    console.log("--- Comparing BOM Sheet ---");
    const wsExp = workbookExport.getWorksheet("BOM");
    const wsTmp = workbookTemplate.getWorksheet("BOM");

    if (!wsExp || !wsTmp) {
        console.log("BOM Sheet missing in one of the files.");
        return;
    }

    // 1. Check Merges
    const mergesExp = Object.keys(wsExp._merges).length;
    const mergesTmp = Object.keys(wsTmp._merges).length;
    console.log(`Merges Count: Export=${mergesExp}, Template=${mergesTmp}`);
    if (mergesExp !== mergesTmp) {
        console.log("WARNING: Merge count mismatch!");
        // console.log("Template Merges:", wsTmp._merges);
    }

    // 2. Check Row 15 Style (Data Row)
    const cellExp = wsExp.getCell("A15");
    const cellTmp = wsTmp.getCell("A15");

    console.log("Export Row 15 Height:", wsExp.getRow(15).height);
    console.log("Template Row 15 Height:", wsTmp.getRow(15).height);

    console.log("\n[A15] Font match?", JSON.stringify(cellExp.font) === JSON.stringify(cellTmp.font));
    if (JSON.stringify(cellExp.font) !== JSON.stringify(cellTmp.font)) {
        console.log("Export Font:", cellExp.font);
        console.log("Template Font:", cellTmp.font);
    }

    console.log("\n[A15] Border match?", JSON.stringify(cellExp.border) === JSON.stringify(cellTmp.border));
    if (JSON.stringify(cellExp.border) !== JSON.stringify(cellTmp.border)) {
        console.log("Export Border:", cellExp.border);
        console.log("Template Border:", cellTmp.border);
    }

    // 3. Check BOQ Sheet
    console.log("\n--- Comparing BOQ Sheet ---");
    const wsBoqExp = workbookExport.getWorksheet("BOQ");
    const wsBoqTmp = workbookTemplate.getWorksheet("BOQ");

    // Row 4 Data
    const cExp = wsBoqExp.getCell("A4");
    const cTmp = wsBoqTmp.getCell("A4");

    console.log("[A4] Font match?", JSON.stringify(cExp.font) === JSON.stringify(cTmp.font));
    if (JSON.stringify(cExp.font) !== JSON.stringify(cTmp.font)) {
        console.log("Export Font:", cExp.font);
        console.log("Template Font:", cTmp.font);
    }
}

compare();
