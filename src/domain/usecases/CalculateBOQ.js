import equipmentData from '../../data/equipment_data.json' with { type: "json" };
import { Material } from '../entities/Material.js';

/**
 * Use Case: Calculate BOQ based on Project Config
 * Implements logic ported from "TOOL_160925_BOQ_BOM_DNO(3 Rail).xlsm" VBA.
 */
export class CalculateBOQ {
    constructor(materialRepository) {
        this.materialRepository = materialRepository;
    }

    async execute(projectConfig) {
        if (!projectConfig.isValid()) {
            throw new Error("Invalid Project Configuration");
        }

        const { name, id: code, dcPower: kwp, pvModel, inverterPosition } = projectConfig;


        // Constants from VBA (Calibrated with DNA2 14.16kWp)
        const MAX_KWP = 500;
        const PV_PER_ARRAY = 8; // User data fits 8 panels/array
        const RAIL_PER_ARRAY = 8; // User data fits 8 rails/array
        const MAX_CABLE_LENGTH = 1000;
        const SITE_COUNT = 1;

        // Accumulators for Cosse (Terminals) - Calculated later based on Inverter


        if (kwp <= 0 || kwp > MAX_KWP) {
            throw new Error(`Invalid Power: ${kwp} kWp (Max ${MAX_KWP})`);
        }

        // 1. Get PV Info
        const pvInfo = this.getPVInfo(pvModel);
        if (!pvInfo) {
            throw new Error(`PV Model '${pvModel}' not found in equipment database.`);
        }
        const pvPowerKw = pvInfo.powerKwp;
        const pvVoc = pvInfo.voc;

        if (pvPowerKw <= 0) throw new Error(`Invalid PV Power for model ${pvModel}`);

        // 2. Calculate Panel Count & Array Stucture
        const pvCount = Math.round(kwp / pvPowerKw);
        if (pvCount < 1) throw new Error("Resulting PV count is less than 1.");

        const numArrays = Math.floor(pvCount / PV_PER_ARRAY) + (pvCount % PV_PER_ARRAY > 0 ? 1 : 0);
        const railCount = (Math.floor(pvCount / PV_PER_ARRAY) * RAIL_PER_ARRAY) + (pvCount % PV_PER_ARRAY > 0 ? RAIL_PER_ARRAY : 0);

        // 3. Materials Calculation (Group I & others)
        // Splice Factor 5 (User: 15 splices / 3 arrays = 5)
        const railConnectors = Math.floor(pvCount / PV_PER_ARRAY) * 5;
        const endClamps = 6 * numArrays;
        const middleClamps = (pvCount * 6 - endClamps) / 2;

        // User Formula Updates:
        // Chân L: =D5*4-3*4 (Rail * 4 - 12)
        // Note: qRail = railCount + 2 (from Line 125 logic). 
        // Let's ensure we use the 'Quantity of Rail' variable correctly. 
        // Line 125 defines qRail. We should probably calculate chanL AFTER qRail is defined or use railCount if that's what D5 represents.
        // Usually D5 in BOM is the FINAL Quantity.
        // So chanL = qRail * 4 - 12.

        // Let's defer exact assignment until variables are clear, or define intermediate.
        // But qRail depends on railCount.

        // Let's redefine these below in the block where I can access qRail if needed?
        // Or just move qRail definition up?

        const qRail = railCount + 2; // Moved up for dependency

        const chanL = (qRail * 4) - 12;
        const kepDay = Math.ceil(pvCount * 1.2); // ROUNDUP(L5*1.2,0)

        const tiepDiaRail = Math.ceil(endClamps / 2); // =G5/2
        const tiepDiaPin = middleClamps + tiepDiaRail; // =H5+K5 (Mid + GroundingRail)

        // 4. Inverter Selection
        const inverterModel = this.selectHuaweiInverter(kwp);
        const invModelUpper = inverterModel.toUpperCase();
        const invInfo = equipmentData.inverters[inverterModel];

        // 5. Cable Lengths (based on Input/Defaults)
        let acCable = projectConfig.acCable || 10; // Default 10m
        let rs485Cable = 10; // Keep default for RS485 for now or add to config if needed
        if (inverterPosition === "Lap xa tu MSB") {
            // Logic from VBA for RS485 if 'Lap xa'
            rs485Cable = 20;
        }
        const dcCable = projectConfig.dcCable;
        const cat5Cable = projectConfig.cat5Cable || 40; // Default 40m

        // 6. MC4 Calculation
        let mc4Count = 0;
        if (invInfo) {
            mc4Count = this.calculateMC4(inverterModel, invInfo.numInputs, pvCount, kwp);
        } else {
            // Fallback if inverter not found in DB
            mc4Count = Math.round(kwp * 0.4);
        }

        // 7. Conduit & Boxes
        // ruột gà lõi thép: =ROUNDUP(S5*0.9,0) -> ROUNDUP(dcCable * 0.9)
        const conduit = Math.ceil(dcCable * 0.9);
        const boxCount = 0; // Logic for boxes was not fully explicit in snippet, assume 0 or calc later

        // --- CONSTRUCT BOQ GROUPS ---

        // Helper to format item
        const mkItem = (group, name, code, unit, qty, notes = "") => ({
            group, name, code, unit, quantity: qty, notes
        });

        const groups = {
            "I": { title: "CLENERGY STRUCTURES", items: [] },
            "II": { title: "PV", items: [] },
            "III": { title: "HUAWEI/CHINA", items: [] }, // Inverter
            "IV": { title: "MC4 Connector, AC Cable, ...", items: [] },
            "V": { title: "He thong tiep dia", items: [] },
            "VI": { title: "Vat tu di day cap AC & DC & LAN", items: [] },
            "VII": { title: "Vat tu chi danh", items: [] },
            "VIII": { title: "Vat tu dau noi", items: [] },
            "IX": { title: "Vat tu phu thi cong lap dat", items: [] },
            "X": { title: "Vat tu tu trung gian va cai tao diem dau noi", items: [] },
            "XI": { title: "Vat tu thi cong bo sung", items: [] },
            "XII": { title: "Vat tu can tai", items: [] },
            "XIII": { title: "Khung lap Tu va Inverter", items: [] }
        };

        // GROUP I
        // groups["I"].items.push(mkItem("I", "PV-ezzRack ECO Rail (5.2m)", "ER-R-ECO/5200", "pcs", Math.ceil((railCount + 2) / 2))); // Duplicate BOM item commented out
        // VBA "GenerateBOMFromBOQ":
        // dataArray(1, 6) = boqMaterials.Item("ER-R-ECO/5200") / 2
        // It seems BOQ Tong stores Total, BOM separates into columns or units?
        // Let's stick to quantities calculated in BOQ Tong logic (Row 866 in reference: dataArray(1, 4) = Rail + 2)

        // Note: qRail defined above now
        const qSplice = railConnectors;
        const qLFoot = chanL;
        const qEndClamp = endClamps;
        const qMidClamp = middleClamps;
        const qCableClip = kepDay;
        const qGroundClip = tiepDiaPin;
        const qGroundLug = tiepDiaRail;

        // GROUP I: Vimetco/ Việt Nam configuration
        groups["I"].title = "Vimetco/ Việt Nam";
        groups["I"].items.push(mkItem("I", "Thanh nhôm 2645 (L=4200mm)", "VMC-SR-2645M-4200", "pcs", qRail));
        groups["I"].items.push(mkItem("I", "Bộ thanh nối Rail 120", "VMC-CL-120-BC", "set", qSplice));
        groups["I"].items.push(mkItem("I", "Chân gà L40xH125xT8mm, bắn vít tôn", "VMC-LKL-40125-8-V6", "set", qLFoot));
        groups["I"].items.push(mkItem("I", "Bộ kẹp biên cao 35mm", "VMC-EC-35", "set", qEndClamp));
        groups["I"].items.push(mkItem("I", "Bộ kẹp giữa đặc 35mm", "VMC-MC-35-D", "set", qMidClamp));
        groups["I"].items.push(mkItem("I", "Kẹp dây cáp cho tấm Pin", "VMC-SCC-SUS", "pcs", qCableClip));
        groups["I"].items.push(mkItem("I", "Lá tiếp địa cho tấm pin-03", "VMC-GCPV-03", "pcs", qGroundClip));
        groups["I"].items.push(mkItem("I", "Bộ Kẹp cáp tiếp địa cho Rail", "VMC-GLPV-SR", "set", qGroundLug));
        groups["I"].items.push(mkItem("I", "Bộ Bulong liên kết M8x25", "VMC-BN-KCL-25", "pcs", 50)); // Updated per request: 50 pcs per site

        // GROUP II: PV + MC4 (as per Template: Panel first, then MC4)
        groups["II"].items.push(mkItem("II", `Tấm pin năng lượng mặt trời ${pvModel}`, pvModel, "panel", pvCount));
        groups["II"].items.push(mkItem("II", "MC4 connector: 4mm2, 1500VDC", "PV-KBT4-EVO", "set", mc4Count));

        // GROUP III: Inverter
        groups["III"].items.push(mkItem("III", `Bộ biến tần Inverter Huawei ${inverterModel}`, inverterModel, "set", 1));
        groups["III"].items.push(mkItem("III", "Smart Dongle-WLAN-FE, WLAN & Fast Ethernet (FE) Communication, Support 3rd Party Monitoring System, IP65 Protection.", "SDongleA-05(AP+STA)", "set", 1));
        // Add Meter (DTSU666-H) - Required for Column P
        groups["III"].items.push(mkItem("III", "Smart Power Sensor, 3 phase 4 wire, Voltage: 230/400V, Current: 250A/50mA (via CT), High accurate: class 1, LCD display, RS485", "DTSU666-H", "set", 1));

        // GROUP IV: Cables (no MC4 here - it's in Group II now)
        groups["IV"].items.push(mkItem("IV", "LV Power Cable, 0.6/1kV, Cu/XLPE/PVC 3x16+1x10sqrt.", "Cu/XLPE/PVC 3x16+1x10sqrt", "Meter", acCable));
        groups["IV"].items.push(mkItem("IV", "Cáp tín hiệu RS485 vặn xoắn chống nhiễu 22AWG, 2 Pair ALTEK KABEL", "Cable RS485 22AWG, 2pair", "Meter", rs485Cable));
        groups["IV"].items.push(mkItem("IV", "Pv Cable Twin Core 4mm2 Solar Cable Connecting Photovoltaic System, 1000VDC, 50A (RED)", "1×4.0", "Meter", dcCable)); // Updated: Full length, no /2
        groups["IV"].items.push(mkItem("IV", "Pv Cable Twin Core 4mm2 Solar Cable Connecting Photovoltaic System, 1000VDC, 50A (BLACK)", "1×4.0", "Meter", dcCable)); // Updated: Full length, no /2
        groups["IV"].items.push(mkItem("IV", "Cáp mạng LS SIMPLE U/UTP, CAT5E, 4 đôi, PVC, 24 AWG, Solid, màu trắng, 305m", "UTP-E-CSG-F1VN-P 0.5X004P/WH", "Meter", cat5Cable)); // Updated: Report in Meter for aggregation

        // GROUP V: Grounding
        // tiếp địa C10: Fixed 10m per site
        groups["V"].items.push(mkItem("V", "PVC INSULATED SINGLE PLEXIBLE WIRE AND CABLE, COPPER CONDUCTOR, 10MM2 - YELLOW-GREEN", "VCm 10MM2 (YELLOW-GREEN)", "Meter", 10));
        groups["V"].items.push(mkItem("V", "Cáp đồng trần C-10mm2 đồng cứng nhiều sợi", "Cu-10mm2", "Meter", Math.ceil(dcCable * 0.8)));
        groups["V"].items.push(mkItem("V", "Cọc đồng tiếp địa D16 - L=1200mm", "", "pcs", 1)); // Updated: 1 pc per site
        groups["V"].items.push(mkItem("V", "Kẹp quả trám cho dây 10mm2 & cọc tiếp địa D16", "", "pcs", 2)); // Updated: 2 pcs per site

        // GROUP VI: Conduit (Detailed below)
        // Generic item removed in favor of OMB12CVL and OMB34CVL in detailed section.

        // GROUP VII & VIII - Add placeholders or logic if specific
        // GROUP VII & VIII
        groups["VII"].items.push(mkItem("VII", "Nap chup dau cap (bo gom 3 cai 3 mau vang - xanh - do) cho day tiet dien 16mm2", "V14", "Set", 10)); // Example quant derived from VBA/10 logic? VBA: 9.25 * Sites. usage 1 site -> ~10
        groups["VII"].items.push(mkItem("VII", "Nap chup dau cap den cho day tiet dien 10mm2", "V8", "Set", 10));
        groups["VII"].items.push(mkItem("VII", "Ong gen co nhiet Haida (bo gom 4 loai 4 mau vang - xanh - do-den) cho day tiet dien 70mm2", "DRS 25", "Meter", 2));
        groups["VII"].items.push(mkItem("VII", "Day thit danh dau day cap 4*100mm 100pcs/pack mau trang (co mieng nhua hinh chu nhat dau day de ghi chu danh)", "", "Pack", 0.1)); // 10 sites = 1 pack
        groups["VII"].items.push(mkItem("VII", "Ong long LM-TU360N (6.0mm, 45m/cal)", "LM-TU360N", "Roll", 0));
        groups["VII"].items.push(mkItem("VII", "Bang in nhan mau vang 9mm Letatwin (MAX) (8m/cuon)", "LM-TP509Y", "pcs", 0));
        groups["VII"].items.push(mkItem("VII", "Muc Den LM-IR300B", "LM-IR300B", "pcs", 0));

        // Group VIII Logic
        let qSC16_8 = 0, qSC10_8 = 0, qSC10_6 = 0, qSC16_6 = 0;
        let qTL25_10 = 0, qTL16_10 = 0, qTL10_10 = 0;

        // Check invModelUpper derived earlier (Group VI logic) or re-derive
        // Re-using invModelUpper from line 216 if available in scope. 
        // Note: invModelUpper is defined inside execute, but need to ensure it's accessible here. 
        // It's defined at line 216. Group VIII is strictly after that.

        if (["SUN2000-8K", "SUN2000-10K", "SUN2000-12K", "SUN2000-15K", "SUN2000-20K"].some(m => invModelUpper.includes(m))) {
            qSC16_8 = 8;
            qSC10_8 = 4;
            qSC10_6 = 4;
        } else if (["SUN2000-30K", "SUN2000-40K", "SUN2000-50K"].some(m => invModelUpper.includes(m))) {
            qTL25_10 = 8;
            qTL16_10 = 4;
            qSC10_6 = 4;
        }

        groups["VIII"].items.push(mkItem("VIII", "Đầu Nối RJ45 DINTEK STP Cat.5e (Chống Nhiễu)", "1501-88054", "pcs", SITE_COUNT * 4));
        groups["VIII"].items.push(mkItem("VIII", "Dau COSSE SC 10-8", "SC 10-8", "pcs", qSC10_8));
        groups["VIII"].items.push(mkItem("VIII", "Dau COSSE SC 16-8", "SC 16-8", "pcs", qSC16_8));
        groups["VIII"].items.push(mkItem("VIII", "Dau COSSE SC 10-6", "SC 10-6", "pcs", qSC10_6));
        groups["VIII"].items.push(mkItem("VIII", "Dau COSSE SC 16-6", "SC 16-6", "pcs", qSC16_6));
        groups["VIII"].items.push(mkItem("VIII", "Dau COSSE TL 25-10", "TL 25-10", "pcs", qTL25_10));
        groups["VIII"].items.push(mkItem("VIII", "Dau COSSE TL 16-10", "TL 16-10", "pcs", qTL16_10));
        groups["VIII"].items.push(mkItem("VIII", "Dau COSSE TL 10-10", "TL 10-10", "pcs", qTL10_10));
        groups["VIII"].items.push(mkItem("VIII", "Dau cosse dong nhom 50mm2", "DTL-2-50-12 MHD", "pcs", Math.round(SITE_COUNT * 1.25)));
        groups["VIII"].items.push(mkItem("VIII", "Dau cosse dong nhom 70mm2", "DTL-2-70-12 MHD", "pcs", Math.round(SITE_COUNT * 0.75)));
        groups["VIII"].items.push(mkItem("VIII", "Bang keo nano mau den", "", "pcs", 0.5)); // 2 sites = 1 roll

        // Update Conduit Fittings based on Logic
        // Logic for AMF (Dau bit) and DNCK (Dau noi)
        let amf12 = 0, amf34 = 0, dnck12 = 0, dnck34 = 0, steelBox = 0;
        if (["SUN2000-8K", "SUN2000-10K", "SUN2000-12K", "SUN2000-15K", "SUN2000-20K"].some(m => invModelUpper.includes(m))) {
            amf12 = 10; amf34 = 5; dnck12 = 12; dnck34 = 14; steelBox = 2;
        } else if (invModelUpper.includes("SUN2000-30K") || invModelUpper.includes("SUN2000-40K")) {
            amf12 = 12; amf34 = 5; dnck12 = 14; dnck34 = 15; steelBox = 3;
        } else if (invModelUpper.includes("SUN2000-50K")) {
            amf12 = 14; amf34 = 5; dnck12 = 18; dnck34 = 17; steelBox = 4;
        }

        groups["VI"].items.push(mkItem("VI", "Ống ruột gà lõi thép bọc nhựa luồn dây điện CVL 1/2\" \r\n(Đường kính trong D15.8-D16.3)", "OMB12CVL", "Meter", conduit - Math.round(conduit * 0.35)));
        groups["VI"].items.push(mkItem("VI", "Dau noi ong ruot ga 1/2 va mang cap dien", "DNCK12", "Pcs", dnck12));
        groups["VI"].items.push(mkItem("VI", "Dau noi ong ruot ga 1/2 va ong ruot ga", "MCK12", "Pcs", Math.round(conduit / 35)));
        groups["VI"].items.push(mkItem("VI", "Dau bit ong ruot ga 1/2", "AMF12", "Pcs", amf12));
        groups["VI"].items.push(mkItem("VI", "Hộp thép HC157D", "MC157D", "Pcs", steelBox));
        groups["VI"].items.push(mkItem("VI", "Nap hop thep NH157", "NH157", "Pcs", steelBox));
        groups["VI"].items.push(mkItem("VI", "Ống ruột gà lõi thép bọc nhựa luồn dây điện CVL 3/4\" \r\n(Đường kính trong D20.7-D21.2)", "OMB34CVL", "Meter", Math.round(conduit * 0.35)));
        groups["VI"].items.push(mkItem("VI", "Dau noi ong ruot ga 3/4 va mang cap", "DNCK34", "Pcs", dnck34));
        groups["VI"].items.push(mkItem("VI", "Dau noi ong ruot ga 3/4 va ong ruot ga", "MCK34", "Pcs", Math.round(conduit * 0.35 / 35)));
        groups["VI"].items.push(mkItem("VI", "Dau bit ong ruot ga 3/4", "AMF34", "Pcs", amf34));
        groups["VI"].items.push(mkItem("VI", "Mang cap hop 100x60mm sử dụng nhựa chống cháy, cây dài 2m, đóng gói theo bó 6 cây", "EH100/60", "Pcs", Math.round(SITE_COUNT * 3.5)));
        groups["VI"].items.push(mkItem("VI", "Ống ruột gà mềm SP D16, 50m/cuộn luồn dây mạng từ router shop đến inverter.", "SP9016CM", "Meter", 50)); // Updated: 50m per site

        // Group IX: Vat tu phu thi cong lap dat (Scaled by kWp)
        groups["IX"].items.push(mkItem("IX", "Sikaflex - 140 Construction (Concrete Grey) 600ml", "", "Tube", Math.ceil(kwp * 0.1)));
        groups["IX"].items.push(mkItem("IX", "Apollo Silicone Sealant A500 300ml", "", "pcs", Math.ceil(kwp * 0.3)));
        groups["IX"].items.push(mkItem("IX", "Keo bọt nở Apollo Foam 750ml", "", "Bottle", Math.ceil(kwp * 0.05)));
        groups["IX"].items.push(mkItem("IX", "Dây rút thép bọc nhựa 7.9x400", "", "Pcs", Math.ceil(kwp * 1.5)));
        groups["IX"].items.push(mkItem("IX", "Ke vuông nhôm định hình (hình ảnh đính kèm)", "", "Pcs", Math.ceil(kwp * 2.2)));
        groups["IX"].items.push(mkItem("IX", "Tắc kê nhựa Fischer M8x50mm cho tường gạch lỗ", "FISCHER-M8x50", "Set", Math.ceil(kwp * 2.5)));
        groups["IX"].items.push(mkItem("IX", "Lông đền phẳng loại dày lỗ 8 cho bulong M8, đường kính ngoài 25mm", "", "Pcs", Math.ceil(kwp * 1.6)));
        groups["IX"].items.push(mkItem("IX", "Vít đầu cờ lê, loại bắt vào tắc kê, M8x50", "", "Pcs", Math.ceil(kwp * 1.9)));
        groups["IX"].items.push(mkItem("IX", "Vít bắn tôn mạ kẽm nhúng nóng dài 6cm \r\n(đuôi cá)", "", "Pcs", Math.ceil(kwp * 2)));
        groups["IX"].items.push(mkItem("IX", "Vít bắn tôn mạ kẽm nhúng nóng dài 10cm \r\n(đuôi cá)", "", "Pcs", Math.ceil(kwp * 3)));
        groups["IX"].items.push(mkItem("IX", "Lồng đền phẳng 6mm", "", "pcs", Math.ceil(kwp * 1.8)));
        groups["IX"].items.push(mkItem("IX", "Ke góc vuông chữ L bản 4cm", "", "Pcs", Math.ceil(kwp * 0.5)));
        groups["IX"].items.push(mkItem("IX", "Dây rút nhựa loại 4x300", "", "Pack", Math.ceil(kwp * 0.05)));
        groups["IX"].items.push(mkItem("IX", "Xi măng - cát lấp hố tiếp địa", "", "Location", Math.ceil(SITE_COUNT / 4)));
        groups["IX"].items.push(mkItem("IX", "Sơn nước màu vàng loại 1kg", "ATM-100", "Can", SITE_COUNT / 3));
        groups["IX"].items.push(mkItem("IX", "Con lăn 6cm", "", "Pcs", SITE_COUNT / 3));

        // Group X: Vat tu tu trung gian va cai tao diem dau noi (Manual Input -> 0)
        // JUNC_BOX Logic

        groups["X"].items.push(mkItem("X", "Tủ sơn tĩnh điện ngoài trời kích thước 400x300x100 dày 1.5mm", "TULE-4030", "pcs", 0));
        groups["X"].items.push(mkItem("X", "Vỏ tủ Suntree SH12PN, ngoài trời", "SH12PN", "set", 0)); // User Req: Default 0
        groups["X"].items.push(mkItem("X", "CABLE GLANDS PG21, GREY, Ø28.3MM", "PG21", "pcs", 0));
        groups["X"].items.push(mkItem("X", "Hàng kẹp lấy mạch áp", "URTK/SS", "pcs", 0));
        groups["X"].items.push(mkItem("X", "Tấm che hàng kẹp", "D-URTK/SS", "pcs", 0));
        groups["X"].items.push(mkItem("X", "PVC INSULATED SINGLE FLEXIBLE WIRE AND CABLE, COPPER CONDUCTOR, 1.5MM2 (30X0.25) - BLACK", "VCm 1.5MM2 (BLACK)", "Meter", 0));
        groups["X"].items.push(mkItem("X", "PVC INSULATED SINGLE FLEXIBLE WIRE AND CABLE, COPPER CONDUCTOR, 25MM2", "VCm 25MM2 (BLACK)", "Meter", 0));
        groups["X"].items.push(mkItem("X", "INSULATED SPADE TERMINALS YF1.5-4 (1PACK = 100PCS) - RED", "YF1.5-4S", "Pack", 0));
        groups["X"].items.push(mkItem("X", "CORD-END TERMINALS (E) RED (1PACK = 100PCS)", "E1510", "Pack", 0));
        groups["X"].items.push(mkItem("X", "KLM-A - Terminal strip marker carrier", "1004348", "pcs", 0));
        groups["X"].items.push(mkItem("X", "Terminal JUT1-35", "CTS35UN", "pcs", 0));
        groups["X"].items.push(mkItem("X", "End Clamp for Terminal Block 35mm2 cable", "CA702", "pcs", 0));
        groups["X"].items.push(mkItem("X", "Partition Plate in Grey colour suitable for CTS35UN Terminal Block", "PP35UN", "pcs", 0));
        groups["X"].items.push(mkItem("X", "DIN RAIL PERFORATED - NS 35/ 7,5 PERF 1000MM - 0807012", "0807012", "Pcs", 0));
        groups["X"].items.push(mkItem("X", "Universal screw type terminal blocks FJ-5N", "FJ-5N", "pcs", 0));
        groups["X"].items.push(mkItem("X", "End Plate Terminal", "FJ-5N", "pcs", 0));
        groups["X"].items.push(mkItem("X", "End Stop", "E/FJ1", "pcs", 0));

        // Group XI: Vat tu thi cong bo sung
        let totalDomino150A = 0, totalDomino200A = 0;
        if (["SUN2000-12K-MAP0", "SUN2000-15KTL-M5", "SUN2000-20KTL-M5"].some(m => invModelUpper.includes(m))) {
            totalDomino150A += 1;
        } else if (["SUN2000-30KTL-M3", "SUN2000-40KTL-M3", "SUN2000-50KTL-M3"].some(m => invModelUpper.includes(m))) {
            totalDomino200A += 1;
        }
        groups["XI"].items.push(mkItem("XI", "Dây móc khóa lò xo co giãn gắn chìa khòa kéo giãn tối thiểu 1.2m", "", "pcs", Math.round(SITE_COUNT)));
        groups["XI"].items.push(mkItem("XI", "Dây xoắn ruột gà bọc dây điện màu đen \r\n(đường kính trong 4mm)", "SWB-06", "Meter", Math.round(SITE_COUNT * 8)));
        groups["XI"].items.push(mkItem("XI", "Mái tôn lạnh  rộng 80cm, dài 50cm cộng vòm 30cm. Dày 4,5 zem phủ màu xanh ngọc", "", "pcs", 0)); // Set to 0 but display
        groups["XI"].items.push(mkItem("XI", "Thanh domino 150A 4P", "SHT-150A-4P", "pcs", totalDomino150A));
        groups["XI"].items.push(mkItem("XI", "Thanh domino 200A 4P", "SHT-200A-4P", "pcs", totalDomino200A));
        groups["XI"].items.push(mkItem("XI", "CB Chống sét Prosurge PV50-1000-V-CD-S 1000VDC 50kA 3P", "PV50-1000-V-CD-S", "set", 2));

        // Group XII: Vat tu can tai
        groups["XII"].items.push(mkItem("XII", "Combo Aptomat + Hộp lắp nổi", "DBSF-63A", "pcs", 1));
        groups["XII"].items.push(mkItem("XII", "Dây điện đơn, đỏ Cadivi 6.0", "CV 6.0", "Meter", 2 * SITE_COUNT));

        // Group XIII: Khung lap Tu va Inverter
        groups["XIII"].items.push(mkItem("XIII", "Tấm Alu che mặt sau của khung thép", "", "pcs", 1));
        groups["XIII"].items.push(mkItem("XIII", "Khung sắt hộp tạo vị trí lắp Tủ và Inverter", "", "pcs", 1));


        const flatItems = [];
        Object.values(groups).forEach(g => flatItems.push(...g.items));

        return {
            projectId: code,
            projectName: name,
            config: projectConfig,
            groups,
            items: flatItems,
            totalCost: 0 // Placeholder - Reading file first
        };
    }

    getPVInfo(model) {
        return equipmentData.photovoltaics[model] || null;
    }

    selectHuaweiInverter(kwp) {
        if (kwp <= 0) return "Unknown";
        if (kwp <= 9.6) return "SUN2000-8K-MAP0";
        if (kwp <= 10.6) return "SUN2000-10K-MAP0";
        if (kwp <= 14.4) return "SUN2000-12K-MAP0";
        if (kwp <= 16) return "SUN2000-15KTL-M5";
        if (kwp <= 25) return "SUN2000-20KTL-M5";
        if (kwp <= 36) return "SUN2000-30KTL-M3";
        if (kwp <= 48) return "SUN2000-40KTL-M3";
        if (kwp <= 60) return "SUN2000-50KTL-M3";
        return "Khong tim thay Inverter phu hop";
    }

    calculateMC4(inverterModel, numInputs, pvCount, kwp) {
        // Updated Logic per User Request (String-based):
        // 12K-20K: 2 Strings -> 2 MC4 + 2 Splices = 4 sets
        // 30K-40K: 3 Strings -> 3 MC4 + 3 Splices = 6 sets
        // 50K: 4 Strings -> 8 sets (extrapolated)

        const invUpper = inverterModel.toUpperCase();

        if (invUpper.includes("8K") || invUpper.includes("10K") || invUpper.includes("12K") || invUpper.includes("15K") || invUpper.includes("20K")) {
            return 4;
        }
        if (invUpper.includes("30K") || invUpper.includes("40K")) {
            return 6;
        }
        if (invUpper.includes("50K")) {
            return 8; // Assumed 4 strings
        }
        if (invUpper.includes("100K")) {
            return 20; // 10 MPPT * 2
        }

        return Math.round(kwp * 0.4); // Fallback
    }
}
