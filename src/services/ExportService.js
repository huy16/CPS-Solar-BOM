
import ExcelJS from 'exceljs';
import * as FileSaver from 'file-saver';
const saveAs = FileSaver.saveAs || FileSaver.default;
import equipmentData from '../data/equipment_data.json' with { type: "json" };

export class ExportService {
    async exportBOQ(reports) {
        try {
            // 1. Load Template
            // Add timestamp to prevent caching of the old template file
            const response = await fetch('/BOQ_BOM_Template.xlsx?v=' + new Date().getTime());
            if (!response.ok) throw new Error("Failed to load template");
            const buffer = await response.arrayBuffer();

            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buffer);

            // 2. Process BOQ Sheet
            const wsBOQ = workbook.getWorksheet("BOQ");
            if (wsBOQ) {
                this.fillBOQSheet(wsBOQ, reports);
            }

            // 3. Process BOM Sheet (If we want to consolidate ALL projects into one BOM)
            // For now, let's aggregate all items from all reports for the BOM
            const wsBOM = workbook.getWorksheet("BOM");
            if (wsBOM) {
                this.fillBOMSheet(wsBOM, reports);
            }

            // 4. Download
            const bufferOut = await workbook.xlsx.writeBuffer();
            const blob = new Blob([bufferOut], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
            saveAs(blob, `BOQ_Export_${new Date().toISOString().slice(0, 10)}.xlsx`);

        } catch (err) {
            console.error("Export failed:", err);
            alert("Export failed: " + err.message);
        }
    }

    fillBOQSheet(ws, reports) {
        // Start from Row 4 (1-based index)
        let currentRow = 4;
        const startDataRow = 4;
        const totalColumns = 23; // After removing DTC and STT columns

        // Capture Style from the first data row (Row 4) to apply to subsequent rows
        const styleRow = ws.getRow(4);
        const rowHeight = styleRow.height || 20; // Capture height
        const styleCache = [];
        for (let i = 1; i <= totalColumns; i++) {
            styleCache[i] = styleRow.getCell(i).style;
        }

        // Define border style
        const borderThin = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };

        reports.forEach((report, index) => {
            const items = report.items;
            const r = ws.getRow(currentRow);
            r.height = rowHeight; // Apply height

            // Apply styles with borders
            for (let i = 1; i <= totalColumns; i++) {
                const cell = r.getCell(i);
                cell.style = { ...styleCache[i], border: borderThin };
            }

            // Column mapping after removing DTC (col 1) and STT (col 2):
            // A (1): MST / Code
            r.getCell(1).value = report.projectId || "";

            // B (2): Project Name  
            r.getCell(2).value = report.projectName || "";

            // --- Map Quantities to Columns ---
            // Helper to sum quantities for items matching a predicate
            const getQty = (predicate) => {
                return items.filter(predicate).reduce((sum, item) => sum + (item.quantity || 0), 0);
            };

            // C (3): Rail
            r.getCell(3).value = getQty(i => i.name.toLowerCase().includes("thanh nhôm") || i.code.includes("2645"));

            // D (4): Nối Rail
            r.getCell(4).value = getQty(i => i.name.toLowerCase().includes("nối rail"));

            // E (5): Chân L
            r.getCell(5).value = getQty(i => i.name.toLowerCase().includes("chân gà") || i.code.includes("LKL"));

            // F (6): Kẹp cuối (End Clamp)
            r.getCell(6).value = getQty(i => i.name.toLowerCase().includes("kẹp biên") || i.name.toLowerCase().includes("kẹp cuối"));

            // G (7): Kẹp giữa (Mid Clamp)
            r.getCell(7).value = getQty(i => i.name.toLowerCase().includes("kẹp giữa"));

            // H (8): Kẹp dây (Cable Clip)
            r.getCell(8).value = getQty(i => i.name.toLowerCase().includes("kẹp dây cáp"));

            // I (9): Tiếp địa PIN (Grounding Clip/Lug for Panel)
            r.getCell(9).value = getQty(i => i.name.toLowerCase().includes("lá tiếp địa")); // VMC-GCPV-03

            // J (10): Tiếp địa Rail (Grounding Lug for Rail)
            r.getCell(10).value = getQty(i => i.name.toLowerCase().includes("kẹp cáp tiếp địa")); // VMC-GLPV-SR

            // K (11): PV
            r.getCell(11).value = getQty(i => i.group === "II");

            // L (12): MC4 (Usually Group IV)
            r.getCell(12).value = getQty(i => i.name.toLowerCase().includes("mc4"));

            // M (13): INVERTER - Show model code instead of quantity
            const inverterItem = items.find(i => i.group === "III" && i.name.toLowerCase().includes("inverter"));
            r.getCell(13).value = inverterItem ? inverterItem.code : "";

            // N (14): SDONGLE
            r.getCell(14).value = getQty(i => i.code === "SdongleA-05");

            // O (15): METER
            r.getCell(15).value = getQty(i => i.code === "DTSU666-H");

            // P (16): CÁP AC
            r.getCell(16).value = getQty(i => i.name.toLowerCase().includes("xlpe") || i.name.toLowerCase().includes("ac cable"));

            // Q (17): TIẾP ĐỊA C10 (Grounding Cable)
            r.getCell(17).value = getQty(i => i.name.includes("C-10mm2") || i.name.includes("VCm 10MM2"));

            // R (18): CÁP DC
            r.getCell(18).value = getQty(i => i.name.includes("PV Cable"));

            // S (19): CÁP MẠNG
            r.getCell(19).value = getQty(i => i.name.toLowerCase().includes("cat5") || i.name.toLowerCase().includes("utp"));

            // T (20): Ruột gà lõi thép (Conduit)
            r.getCell(20).value = getQty(i => i.name.toLowerCase().includes("ruột gà lõi thép") || i.name.includes("OMB"));

            // U (21): JUNC_BOX (Tủ trung gian / Suntree box)
            r.getCell(21).value = getQty(i => i.name.toLowerCase().includes("vỏ tủ suntree") || i.group === "X");

            // V (22): SUNTREE (Maybe the CB/Prosurge set?)
            r.getCell(22).value = getQty(i => i.name.toLowerCase().includes("prosurge") || i.code.includes("PV50"));

            // W (23): NOTE
            r.getCell(23).value = ""; // Leave empty

            r.commit();
            currentRow++;
        });

        // Add SUM row at the end
        const sumRow = ws.getRow(currentRow);
        sumRow.height = rowHeight;

        // Red background style for SUM row
        const sumRowStyle = {
            font: { bold: true, color: { argb: 'FFFFFFFF' } },
            border: borderThin,
            alignment: { horizontal: 'center', vertical: 'middle' },
            fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFF0000' } // Red background
            }
        };

        // First two columns: Label
        sumRow.getCell(1).value = "TỔNG";
        sumRow.getCell(1).style = sumRowStyle;
        sumRow.getCell(2).value = "";
        sumRow.getCell(2).style = sumRowStyle;

        // Add SUM formulas for numeric columns (3 to 22)
        const endDataRow = currentRow - 1;
        for (let col = 3; col <= 22; col++) {
            const colLetter = String.fromCharCode(64 + col); // A=65, so col 3 = C
            sumRow.getCell(col).value = { formula: `SUM(${colLetter}${startDataRow}:${colLetter}${endDataRow})` };
            sumRow.getCell(col).style = sumRowStyle;
        }

        // Last column (NOTE)
        sumRow.getCell(23).value = "";
        sumRow.getCell(23).style = sumRowStyle;

        sumRow.commit();
    }

    fillBOMSheet(ws, reports) {
        // Set column D width to 17.5 as requested
        ws.getColumn(4).width = 17.5;
        // 1. Consolidate items by GROUP
        const groupsData = {};
        reports.forEach(report => {
            if (report.groups) {
                Object.entries(report.groups).forEach(([gid, group]) => {
                    if (!groupsData[gid]) groupsData[gid] = {};
                    group.items.forEach(item => {
                        // FIX: Use name+code as key to prevent merging items with same code (e.g., FJ-5N)
                        const key = `${item.name}|${item.code}`;
                        if (!groupsData[gid][key]) {
                            groupsData[gid][key] = { ...item, quantity: 0 };
                        }
                        groupsData[gid][key].quantity += (item.quantity || 0);
                    });
                });
            }
        });

        // 1b. RE-GROUPING LOGIC (Only for AC Cable now, MC4 is already in Group II)
        Object.keys(groupsData).forEach(gid => {
            const items = groupsData[gid];
            Object.keys(items).forEach(key => {
                const item = items[key];
                const name = item.name ? item.name.toString().toLowerCase() : "";

                // Rule: AC Cable -> Group IV (exclude single wire like 'dây điện đơn')
                if ((name.includes("xlpe") || name.includes("ac cable") || (name.includes("cadivi") && !name.includes("dây điện đơn") && !name.includes("day dien don"))) && !name.includes("pv cable") && !name.includes("dc cable")) {
                    if (gid !== 'IV') {
                        if (!groupsData['IV']) groupsData['IV'] = {};
                        if (groupsData['IV'][key]) groupsData['IV'][key].quantity += item.quantity;
                        else groupsData['IV'][key] = item;
                        delete groupsData[gid][key];
                    }
                }
            });
        });

        // 2. Identify Group Header Rows in Template
        const groupRows = {};
        const romanRegex = /^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII|XIII)$/;
        ws.eachRow((row, rowNumber) => {
            const val = row.getCell(1).value;
            // Match Roman Numeral Strings explicitly
            if (typeof val === 'string' && romanRegex.test(val.trim())) {
                groupRows[val.trim()] = rowNumber;
            }
        });

        // 3. Define Processing Order (Top-Down)
        const processOrder = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII', 'XIII'];

        // Styles
        const fontStyle = { name: 'Cambria', size: 11 };
        const borderThin = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        const borderLeftDouble = { top: { style: 'thin' }, left: { style: 'double' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
        const borderRightDouble = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'double' } };
        const alignCenter = { vertical: 'middle', horizontal: 'center', wrapText: true };
        const alignLeft = { vertical: 'middle', horizontal: 'left', wrapText: true };
        const alignRight = { vertical: 'middle', horizontal: 'right', wrapText: true };

        // Constants
        // const TEMPLATE_ROW_HEIGHT = ws.getRow(15).height || 114; // Removed fixed height

        // Helper for fuzzy matching (Aggressive Normalization)
        const removeVietnameseTones = (str) => {
            str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
            str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
            str = str.replace(/ì|í|ị|ỉ|ĩ/g, "i");
            str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
            str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
            str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
            str = str.replace(/đ/g, "d");
            str = str.replace(/À|Á|Ạ|Ả|Ã|Â|Ầ|Ấ|Ậ|Ẩ|Ẫ|Ă|Ằ|Ắ|Ặ|Ẳ|Ẵ/g, "A");
            str = str.replace(/È|É|Ẹ|Ẻ|Ẽ|Ê|Ề|Ế|Ệ|Ể|Ễ/g, "E");
            str = str.replace(/Ì|Í|Ị|Ỉ|Ĩ/g, "I");
            str = str.replace(/Ò|Ó|Ọ|Ỏ|Õ|Ô|Ồ|Ố|Ộ|Ổ|Ỗ|Ơ|Ờ|Ớ|Ợ|Ở|Ỡ/g, "O");
            str = str.replace(/Ù|Ú|Ụ|Ủ|Ũ|Ư|Ừ|Ứ|Ự|Ử|Ữ/g, "U");
            str = str.replace(/Ỳ|Ý|Ỵ|Ỷ|Ỹ/g, "Y");
            str = str.replace(/Đ/g, "D");
            // Normalize special characters: × -> x, other Unicode symbols
            str = str.replace(/×/g, "x");
            str = str.replace(/÷/g, "/");
            // Normalize: Lowercase -> Remove Punctuation (.,-) -> Remove Spaces
            return str.toLowerCase().replace(/[,.-]/g, "").replace(/\s+/g, "");
        };

        let rowOffset = 0;

        processOrder.forEach((groupId, idx) => {
            if (!groupRows[groupId]) return; // Skip if group not found in template

            const startRow = groupRows[groupId] + rowOffset;
            let endRow;

            // Define default height for this group (try to capture from first data row)
            let defaultGroupHeight = 35; // Fallback
            const potentialRefRow = ws.getRow(startRow + 1);
            if (potentialRefRow && potentialRefRow.height && potentialRefRow.height < 150) {
                defaultGroupHeight = potentialRefRow.height;
            }

            if (groupId === 'XIII') { // XIII (Last Group in template, last in processOrder)
                // Find end of data for the last group
                let r = startRow + 1;
                // Safety limit 200 rows 
                while (r < startRow + 200) {
                    const val = ws.getRow(r).getCell(2).value;
                    const str = val ? val.toString() : "";

                    // Stop at Footer keywords
                    if (str.includes("Hồ sơ") || str.includes("Địa điểm") || str.includes("Ngày") || str.includes("Người lập")) break;

                    // REMOVED: Premature break on empty lines. We want to scan until footer to hiding all empty rows.
                    // if (!val && !ws.getRow(r + 1).getCell(2).value) { break; }
                    r++;
                }
                endRow = r;
            } else {
                // The endRow for the current group is the startRow of the next group physically
                // In Top-Down, next group is idx + 1
                const nextGroupId = processOrder[idx + 1];
                if (groupRows[nextGroupId]) {
                    endRow = groupRows[nextGroupId] + rowOffset;
                } else {
                    endRow = startRow + 2;
                }
            }
            // Prepare Items for this group
            const itemsMap = groupsData[groupId] || {};

            // Extract power from inverter model name (e.g., "SUN2000-20KTL-M5" -> 20)
            const extractPower = (name) => {
                const match = (name || '').match(/(\d+)K/i);
                return match ? parseInt(match[1], 10) : 999;
            };

            // Sort items - special logic for Group III (Inverter)
            let items;
            if (groupId === 'III') {
                // Sort by power capacity for inverters, SDongle at the end
                items = Object.values(itemsMap).sort((a, b) => {
                    const aIsDongle = (a.code || a.name || '').toLowerCase().includes('dongle');
                    const bIsDongle = (b.code || b.name || '').toLowerCase().includes('dongle');
                    if (aIsDongle && !bIsDongle) return 1;
                    if (!aIsDongle && bIsDongle) return -1;
                    return extractPower(a.code || a.name) - extractPower(b.code || b.name);
                });
            } else {
                items = Object.values(itemsMap).sort((a, b) => (a.name || '').localeCompare(b.name || ''));
            }

            // 4. Scan Existing Rows in Template for this Group
            // Range: startRow + 1 to endRow - 1
            // 4. Scan Existing Rows in Template for this Group
            // Range: startRow + 1 to endRow - 1
            const existingRowMap = {};
            const touchedRows = new Set();

            for (let r = startRow + 1; r < endRow; r++) {
                const codeVal = ws.getCell(r, 3).value;
                const nameVal = ws.getCell(r, 2).value;

                // Normalization helper
                const norm = (s) => s ? removeVietnameseTones(s.toString()) : "";

                if (codeVal) {
                    existingRowMap[norm(codeVal)] = r;
                }
                if (nameVal) {
                    const n = norm(nameVal);
                    if (!existingRowMap[n]) existingRowMap[n] = r;
                }
            }

            const itemsToInsert = [];

            // Helper for partial matching
            const findPartialMatch = (itemName) => {
                const normItem = removeVietnameseTones(itemName || "").toLowerCase().substring(0, 30);
                for (const [key, row] of Object.entries(existingRowMap)) {
                    const normKey = key.toLowerCase();
                    if (normKey.includes(normItem) || normItem.includes(normKey.substring(0, 25))) {
                        return row;
                    }
                }
                return null;
            };

            items.forEach(item => {
                const norm = (s) => s ? removeVietnameseTones(s.toString()) : "";
                const code = norm(item.code);
                const name = norm(item.name);

                let targetRow = null;
                // FIX: Prioritize NAME matching first to avoid duplicate code issues (e.g., FJ-5N)
                // 1. Try exact name match FIRST
                if (name && existingRowMap[name] && !touchedRows.has(existingRowMap[name])) {
                    targetRow = existingRowMap[name];
                }
                // 2. Try exact code match (only if not already used)
                else if (code && existingRowMap[code] && !touchedRows.has(existingRowMap[code])) {
                    targetRow = existingRowMap[code];
                }
                // 3. Try partial name match (first 30 chars) - avoid used rows
                else if (name) {
                    const partialMatch = findPartialMatch(item.name);
                    if (partialMatch && !touchedRows.has(partialMatch)) {
                        targetRow = partialMatch;
                    }
                }

                if (targetRow) {
                    // MARK as used immediately to prevent duplicate matching
                    touchedRows.add(targetRow);

                    // UPDATE Existing Row
                    const r = ws.getRow(targetRow);

                    // Force Styles for Consistency

                    // Col 1 (STT): Center + Left Double Border
                    const c1 = r.getCell(1);
                    c1.style = { ...c1.style, border: borderLeftDouble, alignment: alignCenter, font: fontStyle };

                    // Col 2 (Name): Left (Break Shared Style)
                    const c2 = r.getCell(2);
                    c2.style = { ...c2.style, border: borderThin, alignment: { vertical: 'middle', horizontal: 'left', wrapText: true }, font: fontStyle };

                    // Col 3-5 (Code, Supplier, Unit): Center (Break Shared Style) - FORCE ALIGNMENT
                    [3, 4, 5].forEach(colIdx => {
                        const c = r.getCell(colIdx);
                        c.style = { border: borderThin, alignment: alignCenter, font: fontStyle };
                    });

                    // Col 6 (Quantity): Right & Value
                    const c6 = r.getCell(6);
                    c6.value = item.quantity;
                    c6.style = { ...c6.style, border: borderThin, alignment: alignRight };

                    // Col 7-12: Empty borders (ensure thin border is present)
                    for (let c = 7; c <= 12; c++) {
                        const cell = r.getCell(c);
                        cell.style = { ...cell.style, border: borderThin, alignment: alignCenter };
                    }
                    // Col 13: Right Double Border
                    const cell13 = r.getCell(13);
                    cell13.style = { ...cell13.style, border: borderRightDouble, alignment: alignCenter };

                    // Hide row if quantity is 0
                    if (item.quantity === 0 || item.quantity === "0") {
                        r.hidden = true;
                    } else {
                        r.hidden = false; // Ensure visible
                    }
                    // r.commit(); // Ensure changes are saved
                } else {
                    // COLLECT to Insert (only if quantity > 0)
                    if (item.quantity && item.quantity !== 0 && item.quantity !== "0") {
                        itemsToInsert.push(item);
                    }
                }
            });


            // 5. Hide Unused Rows - FIX: Actually hide the row to prevent extra items showing
            for (let r = startRow + 1; r < endRow; r++) {
                if (!touchedRows.has(r)) {
                    // CRITICAL FIX: Hide unused rows to prevent extra items (e.g., 590 when only 715 selected)
                    const row = ws.getRow(r);
                    row.hidden = true;
                    // Also clear the quantity cell to be safe
                    ws.getCell(r, 6).value = "";
                }
            }

            // 5. Insert New Items (if any)
            if (itemsToInsert.length > 0) {
                // Insert AT endRow (pushing the Next Group Header down)
                // Safe fix: insertRows expects array of arrays
                ws.insertRows(endRow, new Array(itemsToInsert.length).fill([]));

                // Fill data and apply styles for NEW items
                itemsToInsert.forEach((item, i) => {
                    const r = ws.getRow(endRow + i);
                    r.height = defaultGroupHeight; // Fix: Use Dynamic Height

                    // Lookup Equipment Data for Rich Description
                    const eCode = item.code ? item.code.toString().trim() : "";
                    const eName = item.name ? item.name.toString().trim() : "";

                    const findInDB = (key) => {
                        if (!key) return null;
                        return equipmentData.inverters[key] || equipmentData.photovoltaics[key];
                    };
                    const dbItem = findInDB(eCode) || findInDB(eName);

                    const finalName = dbItem && dbItem.description ? dbItem.description : item.name;
                    const finalUnit = dbItem && dbItem.unit ? dbItem.unit : item.unit;
                    let finalSupplier = "Vietnam";
                    if (dbItem && dbItem.supplier) {
                        finalSupplier = dbItem.supplier;
                    } else if (['II', 'III'].includes(groupId)) {
                        finalSupplier = "Huawei/China";
                    }

                    const setStyle = (col, border, align) => {
                        const cell = r.getCell(col);
                        // Assign style to avoid shared reference mutation
                        cell.style = {
                            border: border,
                            font: fontStyle,
                            alignment: align
                        };
                    };

                    r.getCell(1).value = "";
                    setStyle(1, borderLeftDouble, alignCenter);

                    r.getCell(2).value = finalName;
                    setStyle(2, borderThin, alignLeft);

                    r.getCell(3).value = item.code;
                    setStyle(3, borderThin, alignCenter);

                    r.getCell(4).value = finalSupplier;
                    setStyle(4, borderThin, alignCenter);

                    r.getCell(5).value = finalUnit;
                    setStyle(5, borderThin, alignCenter);

                    r.getCell(6).value = item.quantity;
                    setStyle(6, borderThin, alignRight);

                    for (let c = 7; c <= 12; c++) {
                        r.getCell(c).value = "";
                        setStyle(c, borderThin, alignCenter);
                    }
                    r.getCell(13).value = "";
                    setStyle(13, borderRightDouble, alignCenter);

                    // DOUBLE FORCE ALIGNMENT (Just to be sure)
                    r.getCell(2).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
                    r.getCell(6).alignment = { vertical: 'middle', horizontal: 'right', wrapText: true };

                    // r.commit();
                });
            }

            // 6. Fix Numbering (STT) for Visible Rows
            let stt = 1;
            const scanLimit = endRow + (itemsToInsert ? itemsToInsert.length : 0);
            for (let r = startRow + 1; r < scanLimit; r++) {
                const row = ws.getRow(r);
                if (!row.hidden && row.getCell(2).value) {
                    const c1 = row.getCell(1);
                    c1.value = stt++;
                    // Force Center Alignment for STT
                    // Explicit style to avoid side-effects
                    c1.style = {
                        ...c1.style,
                        alignment: { vertical: 'middle', horizontal: 'center', wrapText: true },
                        border: { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'double' } } // Double right for col 1? usually thin/double. 
                        // Actually Col 1 boundary is double? Line 446 says borderLeftDouble. 
                        // Let's stick to alignment. Border should be present from default or previous logic.
                        // I will set border just to be safe: Left double, Right thin.
                    };
                    // Re-apply border logic from line 446: setStyle(1, borderLeftDouble, alignCenter)
                    const borderLeftDouble = { top: { style: 'thin' }, left: { style: 'double' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
                    c1.style.border = borderLeftDouble;
                }
            }

            // Accumulate Offset for next groups (Top-Down)
            if (itemsToInsert && itemsToInsert.length > 0) {
                rowOffset += itemsToInsert.length;
            }
        });

        // SPECIAL FIX: Force double border on row 123 + rowOffset (last data row in BOM)
        // This ensures the bottom border matches the double border style of row 13 header
        const lastDataRow = 123 + rowOffset;
        const row123 = ws.getRow(lastDataRow);
        const bottomDoubleBorder = { style: 'double', color: { indexed: 64 } };
        for (let c = 1; c <= 13; c++) {
            const cell = row123.getCell(c);
            const existingBorder = cell.border || {};
            cell.border = {
                ...existingBorder,
                bottom: bottomDoubleBorder
            };
        }

    }
}
