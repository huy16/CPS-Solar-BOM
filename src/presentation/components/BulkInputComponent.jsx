import { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import equipmentData from '../../data/equipment_data.json';
import dashboardBg from '../../assets/dashboard_bg.png';
import solar3dIcon from '../../assets/solar_3d_icon.png';
import React from 'react';

// Sub-component for individual accordion items to maintain state correctly
const AccordionGroup = ({ group, colorMap }) => {
    const [isOpen, setIsOpen] = useState(group.id === "I" || group.id === "II" || group.id === "III");

    return (
        <div className="bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden transition-all duration-300 hover:shadow-md">
            <button
                onClick={() => setIsOpen(!isOpen)}
                className="w-full px-8 py-5 flex items-center justify-between hover:bg-slate-50 transition-colors text-left"
            >
                <div className="flex items-center gap-5">
                    <span className={`w-10 h-10 rounded-xl flex items-center justify-center font-black text-sm italic ${colorMap[group.color]}`}>
                        {group.id}
                    </span>
                    <h4 className="font-bold text-slate-800 text-lg uppercase tracking-tight">{group.title}</h4>
                </div>
                <svg className={`w-6 h-6 text-slate-400 transition-transform duration-300 ${isOpen ? 'rotate-180' : ''}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M19 9l-7 7-7-7" />
                </svg>
            </button>

            <div className={`transition-all duration-500 ease-in-out ${isOpen ? 'max-h-[1200px] border-t border-slate-100 opacity-100' : 'max-h-0 opacity-0 overflow-hidden'}`}>
                <div className="p-0">
                    <table className="w-full text-xs">
                        <thead>
                            <tr className="bg-slate-50/50 text-[10px] font-black text-slate-400 uppercase tracking-[0.2em]">
                                <th className="py-4 px-8 text-left">Vật tư chi tiết</th>
                                <th className="py-4 px-8 text-right">Công thức Engine</th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50">
                            {group.items.map((item, idx) => (
                                <tr key={idx} className="hover:bg-slate-50/30 transition-colors">
                                    <td className="py-4 px-8 text-slate-600 font-medium">{item.n}</td>
                                    <td className="py-4 px-8 text-right">
                                        <code className={`px-3 py-1.5 rounded-lg font-mono text-[11px] font-bold ${colorMap[group.color].replace('bg-', 'bg-opacity-20 bg-')}`}>
                                            {item.f}
                                        </code>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    );
};

export default function BulkInputComponent({ onDataParsed }) {
    const [showFormula, setShowFormula] = useState(false);
    const [previewData, setPreviewData] = useState([]);
    const [activeTab, setActiveTab] = useState("import");
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
    const defaultDC = 80;
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

            if (parsed.length > 0) {
                setPreviewData(prev => [...prev, ...parsed]);
                setActiveTab("list");
            }
        };
        reader.readAsArrayBuffer(file);
        e.target.value = null;
    };

    const handleConfirm = () => {
        onDataParsed(previewData);
        setPreviewData([]);
        setActiveTab("import");
    };

    const calculationGroups = [
        {
            id: "I", title: "Hệ khung (Mounting Structures)", color: "cyan",
            items: [
                { n: "Thanh nhôm 2645 (L4200)", f: "RailCount = (Ceil(Npv / 8) * 8) + 2" },
                { n: "Bộ nối Rail 120", f: "Floor(Npv / 8) * 5" },
                { n: "Chân gà L (Vít tôn)", f: "(RailCount * 4) - 12" },
                { n: "Kẹp biên 35mm", f: "6 * NumArrays (Arrays = Ceil(Npv/8))" },
                { n: "Kẹp giữa", f: "(Npv * 6 - EndClamps) / 2" },
                { n: "Kep day cap pin", f: "Math.ceil(N_pv * 1.2)" },
                { n: "La tiep dia cho pin", f: "MiddleClamps + EndClamps / 2" },
                { n: "Kep tiep dia Rail / Bulong M8", f: "Math.ceil(EndClamps / 2)" }
            ]
        },
        {
            id: "II", title: "Tấm Pin & Đầu nối MC4", color: "energy",
            items: [
                { n: "Tấm pin PV", f: "N_pv = Round(kWp / P_module)" },
                { n: "Bộ đầu nối MC4 1500VDC", f: "Inv 8-20K: 4 bộ | 30-40K: 6 bộ" }
            ]
        },
        {
            id: "III", title: "Biến tần Inverter Huawei", color: "blue",
            items: [
                { n: "Inverter Huawei (8K-50K)", f: "Tự động chọn theo P_dc" },
                { n: "Smart Dongle WLAN-FE", f: "1 bộ / Inverter" }
            ]
        },
        {
            id: "IV", title: "Hệ thống Cáp điện", color: "indigo",
            items: [
                { n: "Cáp AC (Cu/XLPE/PVC)", f: "L_ac (Default 10m)" },
                { n: "Cáp RS485 ALTEK KABEL", f: "10m (Gần) - 20m (Xa)" },
                { n: "Cáp DC 4mm2 (Đỏ/Đen)", f: "L_dc / 2 (m)" },
                { n: "Cáp mạng LAN CAT5E", f: "L_lan (Default 40m)" }
            ]
        },
        {
            id: "V", title: "Hệ thống Tiếp địa", color: "emerald",
            items: [
                { n: "Dây VCm 10mm2 (Vàng-Xanh)", f: "L_dc * 0.8 m" },
                { n: "Cáp đồng trần C10", f: "Site_Count * 10 m" },
                { n: "Cọc tiếp địa D16-L1200", f: "Site_Count * 3 cái" },
                { n: "Kẹp quả trám D16", f: "Site_Count * 3 cái" }
            ]
        },
        {
            id: "VI", title: "Ống luồn & Máng cáp", color: "slate",
            items: [
                { n: "Ống RG lõi thép 1/2\"", f: "Conduit * 0.65 (L_dc * 0.9)" },
                { n: "Ống RG lõi thép 3/4\"", f: "Conduit * 0.35 (L_dc * 0.9)" },
                { n: "Fittings (DNCK/MCK/AMF)", f: "Định mức theo công suất Inv" },
                { n: "Máng cáp hộp 100x60", f: "Round(Site_Count * 3.5) m" }
            ]
        },
        {
            id: "VII", title: "Vật tư chỉ danh (Labeling)", color: "amber",
            items: [
                { n: "Nắp chụp đầu cáp V14/V8", f: "10 bộ / site" },
                { n: "Ống co nhiệt DRS 25", f: "2 m / site" },
                { n: "Dây thít marking 4*100", f: "1 gói / site" },
                { n: "Mực in / Băng nhãn vàng", f: "1 cuộn / site" }
            ]
        },
        {
            id: "VIII", title: "Vật tư đấu nối (Cosse)", color: "orange",
            items: [
                { n: "RJ45 DINTEK Cat.5e", f: "Site_Count * 4 cái" },
                { n: "Cosse SC 10-8 / 16-8", f: "totalSC + Round(Site_Count * 6)" },
                { n: "Cosse Đồng nhôm 50/70mm2", f: "Round(Site_Count * 1.25/0.75)" },
                { n: "Băng keo nano đen", f: "Round(Site_Count * 3)" }
            ]
        },
        {
            id: "IX", title: "Phụ kiện thi công", color: "rose",
            items: [
                { n: "Keo Sikaflex 140 / A500", f: "kWp * 0.1 / kWp * 0.3" },
                { n: "Tắc kê M8x50 / Vít M8", f: "kWp * 2.5 / kWp * 1.9" },
                { n: "Vít đuôi cá 6cm / 10cm", f: "kWp * 2 / kWp * 3" },
                { n: "Dây rút thép 7.9x400", f: "kWp * 1.5" }
            ]
        },
        {
            id: "X", title: "Tủ điện & Thiết bị phụ", color: "purple",
            items: [
                { n: "Vỏ tủ & Thiết bị Group X", f: "0 (Manual Input)" }
            ]
        },
        {
            id: "XI", title: "Thi công hoàn thiện", color: "violet",
            items: [
                { n: "Dây xoắn RG đen SWB-06", f: "Round(Site_Count * 8)" },
                { n: "Mái tôn lạnh bảo vệ", f: "Fixed 1 cái / site" },
                { n: "Chống sét PV50 1000V", f: "Fixed 2 bộ / Project" },
                { n: "Domino 150A/200A", f: "1 cái theo Inverter" }
            ]
        },
        {
            id: "XII", title: "Vật tư cân tải", color: "red",
            items: [
                { n: "Combo At + Hộp nổi 63A", f: "1 bộ / site" },
                { n: "Dây Cadivi 6.0", f: "2 m / site" }
            ]
        },
        {
            id: "XIII", title: "Hệ khung lắp đặt", color: "zinc",
            items: [
                { n: "Tấm Alu che mặt sau", f: "1 tấm / site" },
                { n: "Khung sắt hộp vị trí lắp", f: "1 bộ / site" }
            ]
        }
    ];

    const colorMap = {
        cyan: "bg-cyan-50 border-cyan-100 text-cyan-700",
        energy: "bg-energy-50 border-energy-100 text-energy-700",
        blue: "bg-blue-50 border-blue-100 text-blue-700",
        indigo: "bg-indigo-50 border-indigo-100 text-indigo-700",
        emerald: "bg-emerald-50 border-emerald-100 text-emerald-700",
        slate: "bg-slate-100 border-slate-200 text-slate-700",
        amber: "bg-amber-50 border-amber-100 text-amber-700",
        orange: "bg-orange-50 border-orange-100 text-orange-700",
        rose: "bg-rose-50 border-rose-100 text-rose-700",
        purple: "bg-purple-50 border-purple-100 text-purple-700",
        violet: "bg-violet-50 border-violet-100 text-violet-700",
        red: "bg-red-50 border-red-100 text-red-700",
        zinc: "bg-zinc-100 border-zinc-200 text-zinc-700"
    };

    return (
        <>
            <div className="space-y-3 animate-fade-in-up origin-top transform scale-[0.85]">
                {/* Header Section with Solar Farm Background */}
                <div className="relative rounded-3xl shadow-xl overflow-hidden group min-h-[180px] flex flex-col justify-center items-center text-center p-6 ring-1 ring-slate-900/5">

                    {/* CSS-based Solar Farm Background */}
                    <div className="absolute inset-0 bg-slate-900">
                        <img
                            src={dashboardBg}
                            alt="Solar Farm Background"
                            className="w-full h-full object-fill opacity-95 transition-transform duration-1000 group-hover:scale-105"
                        />
                        <div className="absolute inset-0 bg-gradient-to-t from-slate-900/90 via-slate-900/50 to-slate-900/20 mix-blend-multiply"></div>
                        <div className="absolute inset-0 bg-blue-600/10 mix-blend-overlay"></div>
                    </div>

                    <div className="absolute -top-20 -right-20 w-96 h-96 bg-yellow-500/20 rounded-full blur-3xl animate-pulse-slow"></div>

                    <div className="relative z-10 space-y-2">
                        <h2 className="text-3xl font-black text-white tracking-tight font-display drop-shadow-[0_2px_4px_rgba(0,0,0,0.5)]">
                            Dữ Liệu Dự Án
                        </h2>
                        <p className="text-slate-200 text-lg font-medium mx-auto max-w-2xl bg-slate-900/40 backdrop-blur-md py-1.5 px-6 rounded-full border border-white/10 shadow-lg inline-block">
                            Nhập thông tin các site Solar để tính toán <span className="text-energy-400 font-bold">BOQ & BOM</span> tự động
                        </p>
                    </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
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

                            {/* New Project Status Sidebar Card */}
                            {previewData.length > 0 && (
                                <div className="mt-2 bg-white rounded-xl p-3 border border-slate-200 shadow-sm animate-fade-in-up">
                                    <div className="flex items-center justify-between mb-3">
                                        <span className="text-[10px] font-extrabold text-slate-400 uppercase tracking-wider">Trạng thái Site</span>
                                        <div className="flex h-2 w-2 relative">
                                            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-energy-400 opacity-75"></span>
                                            <span className="relative inline-flex rounded-full h-2 w-2 bg-energy-500"></span>
                                        </div>
                                    </div>

                                    <div className="flex items-center gap-4">
                                        <div className="flex-1 bg-gradient-to-br from-blue-600 to-cyan-500 rounded-lg p-3 text-white shadow-lg shadow-blue-500/20">
                                            <div className="text-[10px] font-bold text-blue-100 uppercase tracking-tighter mb-0.5 opacity-80">Site đã tải lên</div>
                                            <div className="flex items-baseline gap-1.5 focus:outline-none">
                                                <span className="text-2xl font-black font-mono leading-none tracking-tight">{previewData.length}</span>
                                                <span className="text-[10px] font-bold opacity-70">SITES</span>
                                            </div>
                                        </div>

                                        <div className="flex-1 bg-slate-50 border border-slate-100 rounded-lg p-3 border-l-4 border-l-energy-500">
                                            <div className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter mb-0.5">Tổng công suất</div>
                                            <div className="flex items-baseline gap-1 text-slate-800">
                                                <span className="text-lg font-black font-mono leading-none">{previewData.reduce((acc, curr) => acc + curr.dcPower, 0).toFixed(1)}</span>
                                                <span className="text-[9px] font-bold text-slate-400">kWp</span>
                                            </div>
                                        </div>
                                    </div>

                                    <button
                                        onClick={() => setPreviewData([])}
                                        className="w-full mt-3 py-2 text-[10px] font-black text-slate-400 hover:text-red-500 border border-slate-100 hover:border-red-100 rounded-lg flex items-center justify-center gap-2 hover:bg-red-50 transition-all uppercase tracking-widest"
                                    >
                                        <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg>
                                        Xóa tất cả Site
                                    </button>
                                </div>
                            )}
                        </div>

                        <div className="mt-auto pt-4 flex justify-center items-end">
                            <img
                                src={solar3dIcon}
                                alt="Solar 3D Decor"
                                className="w-48 h-auto drop-shadow-2xl hover:scale-105 transition-transform duration-500 filter brightness-110"
                            />
                        </div>
                    </div>

                    <div className="md:col-span-2 bg-white rounded-2xl shadow-sm border border-slate-200 flex flex-col hover:shadow-md transition-shadow duration-300 overflow-hidden">
                        {/* Tab Switcher Headers */}
                        <div className="flex items-center bg-slate-50/80 border-b border-slate-200">
                            <button
                                onClick={() => setActiveTab("import")}
                                className={`flex-1 flex items-center justify-center gap-2 py-4 text-sm font-bold uppercase tracking-tight transition-all relative ${activeTab === "import" ? "text-blue-600 bg-white" : "text-slate-400 hover:text-slate-600"}`}
                            >
                                <svg className={`w-5 h-5 ${activeTab === "import" ? "text-blue-500" : "text-slate-300"}`} fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" /></svg>
                                Nhập Dữ Liệu
                                {activeTab === "import" && <div className="absolute bottom-0 left-0 right-0 h-1 bg-blue-500 rounded-t-full"></div>}
                            </button>
                            <button
                                onClick={() => setActiveTab("list")}
                                className={`flex-1 flex items-center justify-center gap-2 py-4 text-sm font-bold uppercase tracking-tight transition-all relative ${activeTab === "list" ? "text-energy-600 bg-white" : "text-slate-400 hover:text-slate-600"}`}
                            >
                                <svg className={`w-5 h-5 ${activeTab === "list" ? "text-energy-500" : "text-slate-300"}`} fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01" /></svg>
                                Danh sách Site
                                {previewData.length > 0 && (
                                    <span className="inline-flex items-center justify-center w-5 h-5 bg-energy-500 text-white text-[10px] rounded-full scale-75 animate-bounce shadow-md">{previewData.length}</span>
                                )}
                                {activeTab === "list" && <div className="absolute bottom-0 left-0 right-0 h-1 bg-energy-500 rounded-t-full"></div>}
                            </button>
                        </div>

                        {/* Tab Content Area */}
                        <div className="flex-1 p-3 min-h-[400px]">
                            {activeTab === "import" ? (
                                <div className="h-full flex flex-col justify-center items-center gap-5 bg-slate-50/50 rounded-xl border border-dashed border-slate-300 p-6 hover:bg-energy-50/30 hover:border-energy-300 transition-all duration-300">
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

                                        {/* Interactive Steps Guide - Compact */}
                                        <div className="mt-2 w-full grid grid-cols-1 sm:grid-cols-3 gap-2 text-[10px] font-semibold text-slate-500 uppercase tracking-wide">
                                            <div className="bg-white/60 p-2 rounded border border-slate-100 flex items-center justify-center gap-1.5 transition-all hover:bg-white hover:shadow-sm">
                                                <span className="w-4 h-4 rounded-full bg-slate-200 text-slate-600 flex items-center justify-center font-bold text-[9px]">1</span>
                                                <span>Tải Mẫu</span>
                                            </div>
                                            <div className="bg-white/60 p-2 rounded border border-slate-100 flex items-center justify-center gap-1.5 transition-all hover:bg-white hover:shadow-sm">
                                                <span className="w-4 h-4 rounded-full bg-slate-200 text-slate-600 flex items-center justify-center font-bold text-[9px]">2</span>
                                                <span>Nhập Liệu</span>
                                            </div>
                                            <div className="bg-white/60 p-2 rounded border border-slate-100 flex items-center justify-center gap-1.5 transition-all hover:bg-white hover:shadow-sm">
                                                <span className="w-4 h-4 rounded-full bg-slate-200 text-slate-600 flex items-center justify-center font-bold text-[9px]">3</span>
                                                <span>Import</span>
                                            </div>
                                        </div>
                                    </div>

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
                                        <span className="text-sm font-display tracking-wide uppercase font-bold opacity-70">Cơ sở tính toán</span>
                                    </button>
                                </div>
                            ) : (
                                <div className="h-full flex flex-col animate-fade-in-up">
                                    {previewData.length > 0 ? (
                                        <div className="flex-1 flex flex-col border border-slate-100 rounded-xl overflow-hidden bg-slate-50/30 shadow-inner min-h-[380px]">
                                            <div className="px-6 py-3 flex justify-between items-center bg-white border-b border-slate-100">
                                                <h3 className="font-bold text-slate-700 flex items-center gap-2 font-display text-xs uppercase tracking-wide">
                                                    <span className="w-2 h-2 rounded-full bg-energy-500 animate-pulse"></span>
                                                    Danh sách Site đã tải lên
                                                </h3>
                                                <span className="text-[10px] font-black text-slate-400 italic">Scroll to view all rows</span>
                                            </div>

                                            <div className="flex-1 overflow-y-auto custom-scrollbar bg-white max-h-64">
                                                <table className="w-full text-sm">
                                                    <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm border-b border-slate-100">
                                                        <tr className="text-slate-500 text-[10px] font-black uppercase tracking-widest">
                                                            <th className="py-2.5 px-6 text-left">Mã Site</th>
                                                            <th className="py-2.5 px-6 text-left">Tên Dự Án</th>
                                                            <th className="py-2.5 px-6 text-right">kWp</th>
                                                            <th className="py-2.5 px-6 text-center">Ready</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody className="divide-y divide-slate-50">
                                                        {previewData.map((row, i) => (
                                                            <tr key={i} className="hover:bg-energy-50/20 transition-colors group">
                                                                <td className="py-3 px-6 font-mono font-medium text-slate-600 group-hover:text-energy-700">{row.id}</td>
                                                                <td className="py-3 px-6 text-slate-700 font-medium text-xs">{row.name}</td>
                                                                <td className="py-3 px-6 text-right font-mono font-bold text-slate-800">{row.dcPower}</td>
                                                                <td className="py-3 px-6 text-center">
                                                                    <div className="w-1.5 h-1.5 rounded-full bg-energy-500 mx-auto shadow-[0_0_8px_rgba(245,158,11,0.5)]"></div>
                                                                </td>
                                                            </tr>
                                                        ))}
                                                    </tbody>
                                                </table>
                                            </div>

                                            <div className="p-4 bg-white border-t border-slate-100 flex gap-3 mt-auto">
                                                <button
                                                    onClick={() => setActiveTab("import")}
                                                    className="px-6 py-3 bg-slate-100 hover:bg-slate-200 text-slate-600 font-bold rounded-xl transition-all flex items-center gap-2 text-sm uppercase tracking-wide"
                                                >
                                                    Thêm Site
                                                </button>
                                                <button
                                                    onClick={handleConfirm}
                                                    className="flex-1 bg-gradient-to-r from-blue-600 to-energy-600 hover:from-blue-500 hover:to-energy-500 text-white font-black py-3 rounded-xl transition-all shadow-lg shadow-energy-200 hover:shadow-xl hover:shadow-energy-300 hover:-translate-y-0.5 flex items-center justify-center gap-3 text-sm uppercase tracking-widest"
                                                >
                                                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M13 10V3L4 14h7v7l9-11h-7z" />
                                                    </svg>
                                                    Kích Hoạt Tính Toán
                                                </button>
                                            </div>
                                        </div>
                                    ) : (
                                        <div className="h-full flex flex-col justify-center items-center gap-4 text-slate-400 opacity-60 min-h-[380px]">
                                            <svg className="w-16 h-16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
                                            <p className="font-bold uppercase tracking-widest text-xs">Chưa có dữ liệu nào được tải lên</p>
                                            <button onClick={() => setActiveTab("import")} className="text-blue-500 hover:underline font-bold text-sm">Quay lại Nhập Liệu</button>
                                        </div>
                                    )}
                                </div>
                            )}
                        </div>
                    </div>
                </div>

                {/* Removed bottom preview section as it is now in a tab */}
            </div>

            {/* Modal - Outside of the scale transform to prevent clipping */}
            {showFormula && (
                <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-xl" onClick={() => setShowFormula(false)}>
                    <div className="bg-white rounded-[2.5rem] shadow-2xl max-w-5xl w-full max-h-[92vh] flex flex-col border border-white/20 animate-fade-in-up overflow-hidden" onClick={e => e.stopPropagation()}>
                        {/* Modal Header */}
                        <div className="px-10 py-8 bg-slate-900 text-white flex justify-between items-center shrink-0 border-b border-white/10 relative overflow-hidden">
                            <div className="absolute top-0 right-0 w-64 h-64 bg-energy-500/10 blur-[100px] rounded-full -mr-32 -mt-32"></div>
                            <div className="relative z-10 flex items-center gap-6">
                                <div className="w-16 h-16 rounded-2xl bg-gradient-to-tr from-energy-500 to-amber-500 p-0.5 shadow-2xl shadow-energy-500/40">
                                    <div className="w-full h-full rounded-[14px] bg-slate-900 flex items-center justify-center">
                                        <svg className="w-9 h-9 text-energy-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 7h6m0 10v-3m-3 3h.01M9 17h.01M9 14h.01M12 14h.01M15 11h.01M12 11h.01M9 11h.01M7 21h10a2 2 0 002-2V5a2 2 0 00-2-2H7a2 2 0 00-2 2v14a2 2 0 002 2z" />
                                        </svg>
                                    </div>
                                </div>
                                <div>
                                    <h3 className="text-3xl font-black font-display tracking-tight leading-none italic uppercase">Cơ sở tính toán chi tiết</h3>
                                    <p className="text-slate-400 text-xs font-bold uppercase tracking-[0.3em] mt-3 flex items-center gap-2">
                                        <span className="w-2 h-2 rounded-full bg-energy-500 shadow-[0_0_10px_rgba(245,158,11,0.8)] animate-pulse"></span>
                                        Full Engine Specs • Groups I - XIII
                                    </p>
                                </div>
                            </div>
                            <button onClick={() => setShowFormula(false)} className="relative z-10 w-12 h-12 flex items-center justify-center rounded-full bg-white/5 hover:bg-white/20 text-white/40 hover:text-white transition-all group">
                                <svg className="w-8 h-8 transition-transform group-hover:rotate-90" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
                            </button>
                        </div>

                        {/* Modal Content - Interactive Accordion */}
                        <div className="flex-1 overflow-y-auto p-8 space-y-4 custom-scrollbar bg-slate-50/80">
                            {calculationGroups.map((group) => (
                                <AccordionGroup key={group.id} group={group} colorMap={colorMap} />
                            ))}
                        </div>

                        {/* Modal Footer */}
                        <div className="px-10 py-10 bg-slate-900 border-t border-white/5 flex flex-col md:flex-row justify-between items-center shrink-0 gap-6">
                            <div className="flex items-center gap-6">
                                <div className="text-right border-r border-white/10 pr-6">
                                    <span className="block text-[10px] text-slate-500 font-black uppercase tracking-widest leading-none mb-2">Build Signature</span>
                                    <span className="text-white text-xs font-bold font-mono opacity-80">CPS-V3-ALPHA</span>
                                </div>
                                <div className="text-[10px] text-slate-400 font-medium max-w-xs leading-relaxed italic">
                                    * Toàn bộ kết quả tính toán vật tư lẻ (pcs/set/tube) sẽ được tự động làm tròn lên (Ceiling) để đảm bảo đủ vật tư dự phòng trong thi công thực tế tại công trường.
                                </div>
                            </div>
                            <button
                                onClick={() => setShowFormula(false)}
                                className="w-full md:w-auto px-16 py-5 bg-gradient-to-r from-energy-600 to-amber-500 hover:from-energy-500 hover:to-amber-400 text-slate-900 font-black rounded-[1.25rem] transition-all shadow-2xl shadow-energy-500/30 active:scale-95 text-base uppercase tracking-widest"
                            >
                                ĐÃ HIỂU TẤT CẢ
                            </button>
                        </div>
                    </div>
                </div>
            )}
        </>
    );
}
