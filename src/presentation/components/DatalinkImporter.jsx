import { useState } from 'react';
import * as XLSX from 'xlsx';
import { ProjectConfig } from '../../domain/entities/ProjectConfig';

export default function DatalinkImporter({ onDataLoaded }) {
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState(null);

    const handleFileUpload = (e) => {
        const file = e.target.files[0];
        if (!file) return;

        setLoading(true);
        setError(null);

        const reader = new FileReader();
        reader.onload = (evt) => {
            try {
                const bstr = evt.target.result;
                const workbook = XLSX.read(bstr, { type: 'binary' });

                // Find the Datalink sheet (usually has 'DATALINK' in name or is the first one)
                const sheetName = workbook.SheetNames.find(n => n.toUpperCase().includes('DATALINK')) || workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];

                // Parse to JSON
                const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                // Locate Header Row (Look for 'TÊN DỰ ÁN' or 'Project Name')
                let headerRowIndex = rawData.findIndex(row =>
                    row.some(cell => typeof cell === 'string' && (cell.includes('TÊN DỰ ÁN') || cell.includes('PROJECT NAME')))
                );

                if (headerRowIndex === -1) headerRowIndex = 3; // Fallback to standard row 4 (index 3) based on analysis

                const headers = rawData[headerRowIndex];
                const dataRows = rawData.slice(headerRowIndex + 1);

                // Map columns
                // Key Mapping based on analysis: 
                // Index 2: Name, Index 3: Code, Index 4: Panel Count, Index 5: Power
                // We can try to map dynamically or hardcode based on standard template

                const colMap = {
                    name: headers.findIndex(h => h && h.toString().toUpperCase().includes('TÊN DỰ ÁN')),
                    code: headers.findIndex(h => h && (h.toString().toUpperCase().includes('MÃ DỰ ÁN') || h.toString().toUpperCase().includes('MST'))),
                    power: headers.findIndex(h => h && (h.toString().toUpperCase().includes('KWP') || h.toString().toUpperCase().includes('CÔNG SUẤT'))),
                    model: headers.findIndex(h => h && (h.toString().toUpperCase().includes('MODEL') || h.toString().toUpperCase().includes('PV'))),
                };

                const projects = dataRows
                    .map((row, idx) => {
                        const name = row[colMap.name] || `Project ${idx}`;
                        const code = row[colMap.code] || `row-${idx}`;
                        const power = parseFloat(row[colMap.power] || 0);
                        const modelStr = row[colMap.model] || '';

                        if (!power) return null; // Skip if no power

                        // Derive Panel Count from Power / Wattage
                        // 1. Extract wattage from model string (e.g., "555W", "550", "GK715")
                        // Heuristic: Look for number followed by W, or just the largest number in the string roughly between 300-800.
                        let wattage = 580; // Default fallback

                        // Regex for "555W" or "555"
                        const wattMatch = modelStr.toString().match(/(\d{3})\s*W?/i);
                        if (wattMatch && wattMatch[1]) {
                            const parsed = parseInt(wattMatch[1]);
                            if (parsed > 300 && parsed < 800) {
                                wattage = parsed;
                            }
                        }

                        const calculatedPanelCount = Math.ceil((power * 1000) / wattage);

                        return new ProjectConfig(
                            code,
                            name,
                            power,
                            calculatedPanelCount,
                            calculatedPanelCount,
                            '3-Rail',
                            "HSM-ND66-GK715", // Default PV Model
                            "Lap canh tu MSB" // Default Inverter Position
                        );
                    })
                    .filter(Boolean); // Remove nulls

                onDataLoaded(projects);

            } catch (err) {
                setError("Failed to parse Excel file: " + err.message);
                console.error(err);
            } finally {
                setLoading(false);
            }
        };
        reader.readAsBinaryString(file);
    };

    return (
        <div className="p-4 border-2 border-dashed border-gray-300 rounded-lg bg-gray-50 text-center hover:bg-gray-100 transition-colors cursor-pointer relative">
            <input
                type="file"
                accept=".xlsx, .xls"
                onChange={handleFileUpload}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            />
            <div className="flex flex-col items-center justify-center space-y-2">
                <svg className="w-10 h-10 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 13h6m-3-3v6m5 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                <span className="text-gray-600 font-medium">
                    {loading ? 'Parsing...' : 'Click to Upload "Datalink" Excel File'}
                </span>
                <span className="text-xs text-gray-500">Supported formats: .xlsx, .xls</span>
            </div>
            {error && <p className="text-red-500 text-sm mt-2">{error}</p>}
        </div>
    );
}
