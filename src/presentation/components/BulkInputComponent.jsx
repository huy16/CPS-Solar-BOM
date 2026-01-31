import { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import equipmentData from '../../data/equipment_data.json';
import dashboardBg from '../../assets/dashboard_bg.png';
import solar3dIcon from '../../assets/solar_3d_icon.png';

export default function BulkInputComponent({ onDataParsed }) {
    const [showFormula, setShowFormula] = useState(false);
    const [previewData, setPreviewData] = useState([]);
    const fileInputRef = useRef(null);

    // --- PV Model Selection ---
    const pvOptions = Object.keys(equipmentData.photovoltaics || {}).filter(key => {
        const item = equipmentData.photovoltaics[key];
        return item.unit === 'panel';
    });
    const [selectedPV, setSelectedPV] = useState(pvOptions[0] || "HSM-ND66-GK715");

    // --- Hidden Defaults ---
    const defaultPosition = "Lap canh tu MSB";
    const defaultAC = 10;
    const defaultDC = 100;
    const defaultCAT5 = 50;

    // --- File Import Handler ---
    const handleFileUpload = (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            const rows = jsonData.slice(1);

            const parsed = rows.map((row, idx) => {
                if (!row || row.length < 2) return null;

                const code = row[0] ? row[0].toString().trim() : `IMPORT-${idx}`;
                const power = parseFloat(row[1] || 0);
                const name = row[2] ? row[2].toString().trim() : "Shop " + code;

                if (power <= 0) return null;

                return {
                    id: code,
                    name: name,
                    dcPower: power,
                    inverterPosition: defaultPosition,
                    pvModel: selectedPV,
                    acCable: defaultAC,
                    dcCable: defaultDC,
                    cat5Cable: defaultCAT5,
                    panelCount: 0
                };
            }).filter(item => item !== null);

            setPreviewData(prev => [...prev, ...parsed]);
        };
        reader.readAsArrayBuffer(file);
        e.target.value = null;
    };

    const handleConfirm = () => {
        onDataParsed(previewData);
        setPreviewData([]);
    };

    return (
        <div className="space-y-3 animate-fade-in-up">
            {/* Header Section with Solar Farm Background */}
            <div className="relative rounded-3xl shadow-xl overflow-hidden group min-h-[180px] flex flex-col justify-center items-center text-center p-6 ring-1 ring-slate-900/5">

                {/* CSS-based Solar Farm Background */}
                {/* Background Image */}
                <div className="absolute inset-0 bg-slate-900">
                    <img
                        src={dashboardBg}
                        alt="Solar Farm Background"
                        className="w-full h-full object-fill opacity-95 transition-transform duration-1000 group-hover:scale-105"
                    />
                    {/* Gradient Overlay for Depth & Text Readability */}
                    <div className="absolute inset-0 bg-gradient-to-t from-slate-900/90 via-slate-900/50 to-slate-900/20 mix-blend-multiply"></div>
                    <div className="absolute inset-0 bg-blue-600/10 mix-blend-overlay"></div>
                </div>

                {/* Sun Glow Effect */}
                <div className="absolute -top-20 -right-20 w-96 h-96 bg-yellow-500/20 rounded-full blur-3xl animate-pulse-slow"></div>

                <div className="relative z-10 space-y-2">
                    <h2 className="text-3xl font-black text-white tracking-tight font-display drop-shadow-[0_2px_4px_rgba(0,0,0,0.5)]">
                        Dữ Liệu Dự Án
                    </h2>
                    <p className="text-slate-200 text-lg font-medium mx-auto max-w-2xl bg-slate-900/40 backdrop-blur-md py-1.5 px-6 rounded-full border border-white/10 shadow-lg inline-block">
                        Nhập thông tin các trạm Solar để tính toán <span className="text-energy-400 font-bold">BOQ & BOM</span> tự động
                    </p>
                </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                {/* Configuration Card */}
                <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-3 flex flex-col hover:shadow-md transition-shadow duration-300">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2.5 font-display">
                        <span className="w-8 h-8 rounded-lg bg-gradient-to-br from-amber-100 to-orange-100 flex items-center justify-center text-orange-600 shadow-sm ring-1 ring-orange-500/20">
                            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" /><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" /></svg>
                        </span>
                        <span className="bg-gradient-to-r from-slate-800 to-slate-600 bg-clip-text text-transparent">Cấu Hình Chung</span>
                    </h3>

                    <div className="space-y-4 flex-1">
                        <div className="bg-white rounded-xl p-3 border border-slate-200 shadow-sm hover:shadow-md transition-all duration-300">
                            <label className="flex items-center gap-2 text-[10px] font-extrabold text-slate-400 uppercase tracking-wider mb-2">
                                Loại Tấm Pin (PV Module)
                            </label>
                            <div className="relative group">
                                <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400 group-hover:text-energy-500 transition-colors">
                                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 10V3L4 14h7v7l9-11h-7z" /></svg>
                                </div>
                                <select
                                    value={selectedPV}
                                    onChange={(e) => setSelectedPV(e.target.value)}
                                    className="w-full appearance-none bg-slate-50 text-slate-700 text-sm font-bold border border-slate-200 rounded-lg focus:ring-2 focus:ring-energy-500 focus:border-energy-500 py-2.5 pl-10 pr-10 transition-all shadow-sm outline-none cursor-pointer hover:bg-white hover:border-energy-300"
                                >
                                    {pvOptions.map(opt => (
                                        <option key={opt} value={opt}>{opt}</option>
                                    ))}
                                </select>
                                <div className="absolute inset-y-0 right-0 flex items-center pr-3 pointer-events-none text-slate-400 group-hover:text-energy-500">
                                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7" /></svg>
                                </div>
                            </div>

                            {(equipmentData.photovoltaics?.[selectedPV]?.powerKwp) && (
                                <div className="mt-3 relative overflow-hidden bg-gradient-to-br from-amber-50 to-orange-100 rounded-lg p-3 group/cap hover:shadow-md transition-all border border-orange-200/50">
                                    <div className="absolute top-0 right-0 w-16 h-16 bg-gradient-to-br from-orange-400/20 to-transparent rounded-bl-full -mr-4 -mt-4"></div>
                                    <div className="relative flex justify-between items-end">
                                        <div>
                                            <div className="text-[10px] font-bold text-orange-600/80 uppercase tracking-widest mb-0.5">Công suất thiết kế</div>
                                            <div className="font-mono text-xl font-black text-slate-800 tracking-tight flex items-baseline gap-1">
                                                {equipmentData.photovoltaics[selectedPV].powerKwp * 1000}
                                                <span className="text-xs font-bold text-orange-600">Wp</span>
                                            </div>
                                        </div>
                                        <div className="w-8 h-8 rounded-full bg-white/80 flex items-center justify-center border border-orange-200 text-orange-500 shadow-sm group-hover/cap:scale-110 transition-transform">
                                            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 10V3L4 14h7v7l9-11h-7z" /></svg>
                                        </div>
                                    </div>
                                </div>
                            )}
                        </div>
                    </div>

                    {/* Decorative Solar 3D Icon */}
                    <div className="mt-auto pt-4 flex justify-center items-end">
                        <img
                            src={solar3dIcon}
                            alt="Solar 3D Decor"
                            className="w-48 h-auto drop-shadow-2xl hover:scale-105 transition-transform duration-500 filter brightness-110"
                        />
                    </div>
                </div>

                {/* Import Action Card */}
                <div className="md:col-span-2 bg-white rounded-2xl shadow-sm border border-slate-200 p-3 flex flex-col hover:shadow-md transition-shadow duration-300">
                    <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2.5 font-display">
                        <span className="w-8 h-8 rounded-lg bg-gradient-to-br from-blue-100 to-cyan-100 flex items-center justify-center text-blue-600 shadow-sm ring-1 ring-blue-500/20">
                            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
                        </span>
                        <span className="bg-gradient-to-r from-slate-800 to-slate-600 bg-clip-text text-transparent">Nhập Dữ Liệu</span>
                    </h3>

                    <div className="flex-1 flex flex-col justify-center items-center gap-5 bg-slate-50/50 rounded-xl border border-dashed border-slate-300 p-6 hover:bg-energy-50/30 hover:border-energy-300 transition-all duration-300">
                        {/* Vertical Button Stack */}
                        <div className="flex flex-col w-full max-w-sm gap-3">
                            <a
                                href="/Input_Template.xlsx"
                                download
                                className="group w-full inline-flex items-center justify-center gap-2 bg-gradient-to-r from-emerald-500 to-teal-600 text-white text-sm font-bold px-6 py-3 rounded-xl hover:from-emerald-400 hover:to-teal-500 hover:shadow-lg hover:shadow-emerald-500/40 hover:-translate-y-0.5 transition-all shadow-md ring-1 ring-white/20"
                            >
                                <svg className="w-5 h-5 text-emerald-100 group-hover:text-white transition-colors" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                                </svg>
                                1. Tải File Mẫu (Input_Template.xlsx)
                            </a>

                            <div className="relative w-full">
                                <input
                                    type="file"
                                    accept=".xlsx,.xls"
                                    ref={fileInputRef}
                                    onChange={handleFileUpload}
                                    className="hidden"
                                />
                                <button
                                    onClick={() => fileInputRef.current?.click()}
                                    className="w-full inline-flex items-center justify-center gap-2 bg-gradient-to-r from-cyan-600 to-blue-600 text-white text-sm font-bold px-6 py-3 rounded-xl hover:from-cyan-500 hover:to-blue-500 hover:shadow-lg hover:shadow-cyan-500/50 hover:-translate-y-0.5 transition-all shadow-md ring-1 ring-white/20"
                                >
                                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                                    </svg>
                                    2. Upload File Excel (Import Data)
                                </button>
                            </div>
                        </div>

                        {/* Interactive Steps Guide - Compact */}
                        <div className="mt-2 w-full grid grid-cols-1 sm:grid-cols-3 gap-2 text-[10px] font-semibold text-slate-500 uppercase tracking-wide">
                            <div className="bg-white/60 p-2 rounded border border-slate-100 flex items-center justify-center gap-1.5">
                                <span className="w-4 h-4 rounded-full bg-slate-200 text-slate-600 flex items-center justify-center font-bold text-[9px]">1</span>
                                <span>Tải Mẫu</span>
                            </div>
                            <div className="bg-white/60 p-2 rounded border border-slate-100 flex items-center justify-center gap-1.5">
                                <span className="w-4 h-4 rounded-full bg-slate-200 text-slate-600 flex items-center justify-center font-bold text-[9px]">2</span>
                                <span>Nhập Liệu</span>
                            </div>
                            <div className="bg-white/60 p-2 rounded border border-slate-100 flex items-center justify-center gap-1.5">
                                <span className="w-4 h-4 rounded-full bg-slate-200 text-slate-600 flex items-center justify-center font-bold text-[9px]">3</span>
                                <span>Import</span>
                            </div>
                        </div>

                        {/* Calculation Formula Button */}
                        <button
                            onClick={() => setShowFormula(true)}
                            className="w-full max-w-sm mt-5 py-2.5 px-4 bg-slate-50 hover:bg-white text-slate-600 hover:text-energy-600 font-medium rounded-xl border border-dashed border-slate-300 hover:border-energy-300 flex items-center justify-center gap-2 transition-all duration-300 group"
                            title="Xem cơ sở tính toán BOM"
                        >
                            <span className="w-8 h-8 rounded-lg bg-white border border-slate-200 text-slate-400 group-hover:text-energy-500 group-hover:border-energy-200 flex items-center justify-center transition-colors">
                                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707M16 12a4 4 0 11-8 0 4 4 0 018 0z" />
                                </svg>
                            </span>
                            <span className="text-sm font-display tracking-wide">Cơ sở tính toán</span>
                        </button>
                        <p className="mt-2 text-[10px] text-center text-slate-400 font-medium italic max-w-sm mx-auto flex items-center justify-center gap-1.5">
                            📌 <span>Số liệu dùng để tham khảo, cần xác nhận từ kỹ sư thiết kế trước khi sử dụng</span>
                        </p>
                    </div>
                </div>
            </div>

            {/* Preview Table Section */}
            {previewData.length > 0 && (
                <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden animate-fade-in-up">
                    <div className="px-6 py-4 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                        <h3 className="font-bold text-slate-700 flex items-center gap-3 font-display">
                            <span className="w-8 h-8 rounded-lg bg-energy-100 text-energy-700 flex items-center justify-center text-sm font-extrabold font-mono">{previewData.length}</span>
                            Dự án đã tải lên
                        </h3>
                        <div className="flex gap-2">
                            <button
                                onClick={() => setPreviewData([])}
                                className="text-xs text-red-500 hover:text-red-700 font-bold px-3 py-1.5 hover:bg-red-50 rounded-lg transition-colors uppercase tracking-wide"
                            >
                                Xóa tất cả
                            </button>
                        </div>
                    </div>

                    <div className="max-h-80 overflow-y-auto custom-scrollbar">
                        <table className="w-full text-sm">
                            <thead className="bg-white sticky top-0 z-10 shadow-sm border-b border-slate-100">
                                <tr className="text-slate-500 text-xs font-bold uppercase tracking-wider">
                                    <th className="py-3 px-6 text-left bg-slate-50/80 backdrop-blur">Mã Trạm</th>
                                    <th className="py-3 px-6 text-left bg-slate-50/80 backdrop-blur">Tên Dự Án</th>
                                    <th className="py-3 px-6 text-right bg-slate-50/80 backdrop-blur">Công Suất (kWp)</th>
                                    <th className="py-3 px-6 text-center bg-slate-50/80 backdrop-blur">Trạng Thái</th>
                                </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-50">
                                {previewData.map((row, i) => (
                                    <tr key={i} className="hover:bg-energy-50/30 transition-colors group">
                                        <td className="py-3 px-6 font-mono font-medium text-slate-600 group-hover:text-energy-700 transition-colors">{row.id}</td>
                                        <td className="py-3 px-6 text-slate-700 font-medium">{row.name}</td>
                                        <td className="py-3 px-6 text-right font-mono font-bold text-slate-800">{row.dcPower}</td>
                                        <td className="py-3 px-6 text-center">
                                            <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider bg-energy-100 text-energy-700 border border-energy-200">
                                                Ready
                                            </span>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>

                    <div className="p-6 bg-slate-50/50 border-t border-slate-100">
                        <button
                            onClick={handleConfirm}
                            className="w-full bg-gradient-to-r from-cyan-600 to-energy-600 hover:from-cyan-500 hover:to-energy-500 text-white font-bold py-4 rounded-xl transition-all shadow-lg shadow-energy-200 hover:shadow-xl hover:shadow-energy-300 hover:-translate-y-0.5 flex items-center justify-center gap-3 text-lg font-display tracking-tight"
                        >
                            <svg className="w-6 h-6 animate-pulse" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 10V3L4 14h7v7l9-11h-7z" />
                            </svg>
                            Kích Hoạt Tính Toán ({previewData.length} Dự Án)
                        </button>
                    </div>
                </div>
            )}

            {/* Formula Modal */}
            {showFormula && (
                <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm" onClick={() => setShowFormula(false)}>
                    <div className="bg-white rounded-2xl shadow-2xl max-w-lg w-full p-6 border border-slate-200 animate-fade-in-up" onClick={e => e.stopPropagation()}>
                        <div className="flex justify-between items-center mb-4">
                            <h3 className="text-xl font-bold text-slate-800 font-display flex items-center gap-2">
                                <span className="w-8 h-8 rounded-lg bg-energy-100/50 flex items-center justify-center text-energy-600">
                                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 7h6m0 10v-3m-3 3h.01M9 17h.01M9 14h.01M12 14h.01M15 11h.01M12 11h.01M9 11h.01M7 21h10a2 2 0 002-2V5a2 2 0 00-2-2H7a2 2 0 00-2 2v14a2 2 0 002 2z" /></svg>
                                </span>
                                Cơ sở tính toán BOM
                            </h3>
                            <button onClick={() => setShowFormula(false)} className="text-slate-400 hover:text-slate-600 transition-colors">
                                <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
                            </button>
                        </div>

                        <div className="space-y-4">
                            <div className="bg-slate-50 p-4 rounded-xl border border-slate-200">
                                <div className="text-sm text-slate-500 mb-2 font-medium uppercase tracking-wider">Công thức chính</div>
                                <div className="font-mono text-lg font-bold text-slate-800 text-center py-2 bg-white rounded-lg border border-slate-100 shadow-sm">
                                    P<sub>dc</sub> = N<sub>panels</sub> × P<sub>module</sub>
                                </div>
                            </div>

                            <div className="space-y-2">
                                <div className="flex items-start gap-3 text-sm text-slate-600">
                                    <span className="font-mono font-bold text-energy-600 w-16">P<sub>dc</sub></span>
                                    <span>Tổng công suất DC thiết kế (Wp/kWp)</span>
                                </div>
                                <div className="flex items-start gap-3 text-sm text-slate-600">
                                    <span className="font-mono font-bold text-energy-600 w-16">N<sub>panels</sub></span>
                                    <span>Số lượng tấm pin (modules)</span>
                                </div>
                                <div className="flex items-start gap-3 text-sm text-slate-600">
                                    <span className="font-mono font-bold text-energy-600 w-16">P<sub>module</sub></span>
                                    <span>Công suất danh định của 1 tấm pin (Wp)</span>
                                </div>
                            </div>

                            <div className="pt-4 border-t border-slate-100 flex justify-end">
                                <button onClick={() => setShowFormula(false)} className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium rounded-lg transition-colors">Đóng</button>
                            </div>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}
