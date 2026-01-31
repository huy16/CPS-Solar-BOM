
import { CalculateBOQ } from '../src/domain/usecases/CalculateBOQ.js';
import { ExportService } from '../src/services/ExportService.js';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';

// Mock Config
const mockConfig = {
    isValid: () => true,
    name: "Dự án Mẫu Agent",
    id: "AGENT-007",
    dcPower: 50, // 50 kWp
    pvModel: "SUN2000-50KTL-M3", // Wait, this is inverter. PV model?
    // CalculateBOQ uses pvModel for panels.
    // Let's us a valid panel model from DB or mapped string
    // In CalculateBOQ: const pvInfo = this.getPVInfo(pvModel);
    // Let's use a key that exists in equipment_data.json if possible.
    // Or just a string if it's tolerant.
    // In equipment_data.json lines 292: getPVInfo returns equipmentData.photovoltaics[model]
    // I need a valid key. I'll check equipment data or guess.
    // Defaulting to "Jinko 550W" usually works if seeded.
    // Let's assume the user has data. If not, it throws.
    // I'll define a mock pvModel that works or mock the valid function.
    pvModel: "JKM550M-72HL4-V",
    inverterPosition: "Lap gan MSB",
    acCable: 50,
    dcCable: 200,
    cat5Cable: 100
};

// Mock Repo
const mockRepo = { find: () => null };

(async () => {
    try {
        console.log("Loading Template...");
        const templatePath = path.resolve('public/BOQ_BOM_Template.xlsx');
        if (!fs.existsSync(templatePath)) {
            throw new Error(`Template not found at ${templatePath}`);
        }

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(templatePath);
        const wsBOM = wb.getWorksheet('BOM'); // or .getWorksheet(2)

        if (!wsBOM) throw new Error("Sheet BOM not found");

        console.log("Calculating BOQ...");
        const calculator = new CalculateBOQ(mockRepo);

        // Monkey-patch getPVInfo to avoid DB issues if key missing
        calculator.getPVInfo = () => ({ powerKwp: 0.55, voc: 50 }); // 550W panel

        const boqData = await calculator.execute(mockConfig);
        console.log("BOQ Data Ready.");

        console.log("Filling BOM Sheet...");
        const svc = new ExportService();
        svc.fillBOMSheet(wsBOM, [boqData]);

        const outFile = path.resolve('Sample_BOM_Agent_Verified.xlsx');
        await wb.xlsx.writeFile(outFile);
        console.log(`SUCCESS: Created ${outFile}`);

    } catch (e) {
        console.error("FAIL:", e);
    }
})();
