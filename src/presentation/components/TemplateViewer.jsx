import { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

const TemplateViewer = () => {
    const [workbook, setWorkbook] = useState(null);
    const [selectedSheet, setSelectedSheet] = useState(null);
    const [sheetData, setSheetData] = useState([]);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState(null);

    useEffect(() => {
        const fetchTemplate = async () => {
            try {
                const response = await fetch('/BOQ_BOM_Template.xlsx');
                if (!response.ok) throw new Error('Failed to fetch template file');

                const blob = await response.blob();
                const arrayBuffer = await blob.arrayBuffer();
                const wb = XLSX.read(arrayBuffer);

                setWorkbook(wb);
                if (wb.SheetNames.length > 0) {
                    selectSheet(wb, wb.SheetNames[0]);
                }
                setLoading(false);
            } catch (err) {
                console.error("Error loading template:", err);
                setError(err.message);
                setLoading(false);
            }
        };

        fetchTemplate();
    }, []);

    const selectSheet = (wb, sheetName) => {
        const ws = wb.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        setSelectedSheet(sheetName);
        setSheetData(data);
    };

    if (loading) return <div className="p-8 text-center">Loading Template...</div>;
    if (error) return <div className="p-8 text-center text-red-500">Error: {error}</div>;

    return (
        <div className="bg-white rounded-lg shadow-sm p-6">
            <h2 className="text-xl font-bold mb-4">BOQ & BOM Template Viewer</h2>

            {/* Sheet Tabs & Download */}
            <div className="flex justify-between items-center border-b mb-4 pb-2">
                <div className="flex gap-2">
                    {workbook?.SheetNames.map(name => (
                        <button
                            key={name}
                            onClick={() => selectSheet(workbook, name)}
                            className={`px-4 py-2 font-medium rounded-t-lg transition-colors ${selectedSheet === name
                                    ? 'bg-blue-50 text-blue-600 border-b-2 border-blue-600'
                                    : 'text-gray-500 hover:text-gray-700 hover:bg-gray-50'
                                }`}
                        >
                            {name}
                        </button>
                    ))}
                </div>
                <a
                    href="/BOQ_BOM_Template.xlsx"
                    download="BOQ_BOM_Template.xlsx"
                    className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors text-sm font-medium"
                >
                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4-4m0 0l-4 4m4-4v12" /></svg>
                    Download Tool Excel
                </a>
            </div>

            {/* Data Table */}
            <div className="overflow-x-auto">
                <table className="min-w-full divide-y divide-gray-200 border">
                    <tbody className="bg-white divide-y divide-gray-200">
                        {sheetData.map((row, rowIndex) => (
                            <tr key={rowIndex} className={rowIndex === 0 ? 'bg-gray-50' : ''}>
                                {row.map((cell, cellIndex) => (
                                    <td
                                        key={cellIndex}
                                        className={`px-3 py-2 text-sm border-r ${rowIndex === 0 ? 'font-bold text-gray-900' : 'text-gray-700'
                                            }`}
                                    >
                                        {cell}
                                    </td>
                                ))}
                            </tr>
                        ))}
                    </tbody>
                </table>
                {sheetData.length === 0 && <p className="p-4 text-gray-500">Empty Sheet</p>}
            </div>
        </div>
    );
};

export default TemplateViewer;
