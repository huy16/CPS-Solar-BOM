import * as XLSX from 'xlsx';

// Create workbook
const wb = XLSX.utils.book_new();

// ==========================================
// 1. BOQ SHEET
// ==========================================
// Row 1: "2. BILL OF QUANTITIES FOR INSTALLATION"
// Row 2: Headers (Groups)
// Row 3: Sub-Headers (Items)

const boqData = [
    ["2. BILL OF QUANTITIES FOR INSTALLATION"],
    ["ĐTC", "STT", "MST", "TÊN DỰ ÁN", "CLENERGY STRUCTURES", null, null, null, null, null, null, null, null, null, "HUAWEI", null, null, "CÁP", null, null, null, null, "JUNC_BOX", "SUNTREE", "NOTE"],
    [null, null, null, null, "Rail", "Nối Rail", "Chân L", "Kẹp cuối", "Kẹp giữa", "Kẹp dây", "Tiếp địa PIN", "Tiếp địa rail", "PV", "MC4", "INVERTER", "SDONGLE", "METER", "CÁP AC", "TIẾP ĐỊA C10", "CÁP DC", "CÁP MẠNG", "Ruột gà lõi thép", null, null, null]
];

const wsBOQ = XLSX.utils.aoa_to_sheet(boqData);

// MERGES for BOQ
wsBOQ['!merges'] = [
    // Title
    { s: { r: 0, c: 0 }, e: { r: 0, c: 24 } },
    // Group Headers
    { s: { r: 1, c: 4 }, e: { r: 1, c: 13 } }, // CLENERGY STRUCTURES (E2:N2)
    { s: { r: 1, c: 14 }, e: { r: 1, c: 16 } }, // HUAWEI (O2:Q2)
    { s: { r: 1, c: 17 }, e: { r: 1, c: 21 } }, // CÁP (R2:V2)
    // Vertical Merges for Main Headers (A2:A3, etc.)
    { s: { r: 1, c: 0 }, e: { r: 2, c: 0 } }, // ĐTC
    { s: { r: 1, c: 1 }, e: { r: 2, c: 1 } }, // STT
    { s: { r: 1, c: 2 }, e: { r: 2, c: 2 } }, // MST
    { s: { r: 1, c: 3 }, e: { r: 2, c: 3 } }, // TÊN DỰ ÁN
    { s: { r: 1, c: 22 }, e: { r: 2, c: 22 } }, // JUNC_BOX
    { s: { r: 1, c: 23 }, e: { r: 2, c: 23 } }, // SUNTREE
    { s: { r: 1, c: 24 }, e: { r: 2, c: 24 } }  // NOTE
];

// Column Widths
wsBOQ['!cols'] = [
    { wch: 8 },  // ĐTC
    { wch: 5 },  // STT
    { wch: 10 }, // MST
    { wch: 30 }, // TÊN DỰ ÁN
];
for (let i = 4; i <= 24; i++) wsBOQ['!cols'].push({ wch: 12 }); // Default for others

XLSX.utils.book_append_sheet(wb, wsBOQ, "BOQ");

// ==========================================
// 2. BOM SHEET
// ==========================================
// Matches strict format from "Template BOM&BOQ.xlsx"

const bomData = [];
// Empty rows 1-5
for (let i = 0; i < 5; i++) bomData.push([]);

// Row 6: Title
bomData.push(["PHIẾU LIỆT KÊ YÊU CẦU VẬT TƯ - DỊCH VỤ"]);

// Row 7: Empty
bomData.push([]);

// Row 8: Info 1
bomData.push([null, "- Số phiếu: ", "CAS-MWG_1.0", null, null, null, null, null, "- Mã tủ: "]);

// Row 9: Info 2
bomData.push([null, "- Khách hàng:", "MWG", null, null, null, null, null, "- Mã dự án: CAS-MWG-171"]);

// Row 10: Info 3
bomData.push([null, "- Hợp đồng số: ", null, null, null, null, null, null, "- Hạng mục: Vật tư thi công 8 shop "]);

// Row 11: Info 4
bomData.push([null, "- Lý do mua hàng:", "Vật tư thi công tại công trường tại Đà Nẵng"]);

// Row 12: Empty
bomData.push([]);

// Row 13: Headers
bomData.push([
    "TT",
    "Tên hàng\n(quy cách, chủng loại)",
    "Mã hàng",
    "Nhà CC./ SX",
    "ĐVT",
    "Thiết kế\n(1)",
    "Dự phòng\n(2)",
    "Yêu cầu\n(3)=(1)+(2)",
    "Tồn kho\n(4)",
    "Đặt hàng\n(5)=(3)-(4)",
    "Ngày cần hàng",
    "Ngày dự kiến\n có hàng",
    "Ghi\n chú"
]);

const wsBOM = XLSX.utils.aoa_to_sheet(bomData);

// MERGES for BOM
wsBOM['!merges'] = [
    { s: { r: 5, c: 0 }, e: { r: 5, c: 12 } } // Title Row 6
];

// Column Widths
wsBOM['!cols'] = [
    { wch: 5 },  // TT
    { wch: 40 }, // Tên hàng
    { wch: 20 }, // Mã hàng
    { wch: 20 }, // Nhà CC
    { wch: 8 },  // ĐVT
    { wch: 10 }, // Thiết kế
    { wch: 10 }, // Dự phòng
    { wch: 12 }, // Yêu cầu
    { wch: 10 }, // Tồn kho
    { wch: 12 }, // Đặt hàng
    { wch: 12 }, // Ngày cần
    { wch: 12 }, // Ngày có
    { wch: 20 }  // Ghi chú
];

XLSX.utils.book_append_sheet(wb, wsBOM, "BOM");

// Write File
const fileName = "BOQ_BOM_Template.xlsx";
XLSX.writeFile(wb, fileName);
console.log(`Successfully created '${fileName}' with extracted structure.`);
