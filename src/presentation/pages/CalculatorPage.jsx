
import { useState } from 'react';
import { ProjectConfig } from '../../domain/entities/ProjectConfig';
import { DI } from '../../infrastructure/di/Container';
import DatalinkImporter from '../components/DatalinkImporter';
import BulkInputComponent from '../components/BulkInputComponent';
import WorkflowStepper from '../components/WorkflowStepper';
import equipmentData from '../../data/equipment_data.json';
import { ExportService } from '../../services/ExportService';

export default function CalculatorPage() {
    // Workflow State
    const [activeStep, setActiveStep] = useState(0); // 0: Input, 1: Result (Skipping explicit Review for now as it's built-in to components)
    // Actually, let's Stick to the plan:
    // Step 0: Input Selection & Data Entry
    // Step 1: Verification/Preview (Implicit in Smart Paste, Explicit for File?) -> Let's keep it simple first
    // Let's define steps:
    const STEPS = [
        { label: 'Input Data' },
        { label: 'Calculation Results' }
    ];

    // Mode State (within Step 0)
    const [inputMode, setInputMode] = useState(null); // 'single', 'paste', 'file' or null for "Landing"

    // Data State
    const [projectsToCalculate, setProjectsToCalculate] = useState([]);

    // Result State
    const [reports, setReports] = useState([]);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState(null);

    // UI State for Results
    const [selectedShops, setSelectedShops] = useState(new Set());
    const [expandedShops, setExpandedShops] = useState(new Set());

    // Single Form State
    const [singleFormData, setSingleFormData] = useState({
        projectName: 'Solar Park A',
        dcPower: 100,
        panelCount: 140, // 100 / 0.715 ~= 140
        pvModel: "HSM-ND66-GK715",
        inverterPosition: "Lap canh tu MSB"
    });

    const resetWorkflow = () => {
        setActiveStep(0);
        setInputMode(null);
        setProjectsToCalculate([]);
        setReports([]);
        setError(null);
    };

    const handleSingleChange = (e) => {
        const { name, value } = e.target;
        setSingleFormData(prev => ({ ...prev, [name]: value }));
    };

    const runCalculation = async (projects) => {
        setLoading(true);
        setError(null);
        try {
            const calculateUseCase = DI.resolveCalculateBOQ();
            // Allow ui to verify before this? 
            // For now, straight to calculation as requested by "Click Calculate" flow
            const results = await Promise.all(projects.map(p => calculateUseCase.execute(p)));
            setReports(results);
            setActiveStep(1); // Move to Result Step
        } catch (err) {
            setError("Calculation Error: " + err.message);
        } finally {
            setLoading(false);
        }
    };

    // --- Handlers for Input Modes ---

    const handleSingleSubmit = async () => {
        const config = new ProjectConfig(
            Date.now().toString(), // Temp ID
            singleFormData.projectName,
            Number(singleFormData.dcPower),
            Number(singleFormData.panelCount),
            '3-Rail',
            singleFormData.pvModel,
            singleFormData.inverterPosition
        );
        await runCalculation([config]);
    };

    const handlePasteSubmit = async (parsedData) => {
        const projects = parsedData.map(d => new ProjectConfig(
            d.id,
            d.name,
            d.dcPower,
            d.panelCount,
            '3-Rail',
            d.pvModel,
            d.inverterPosition,
            d.acCable,
            d.dcCable,
            d.cat5Cable
        ));
        await runCalculation(projects);
    };

    const handleFileImported = async (projects) => {
        await runCalculation(projects);
    };

    const handleExport = async () => {
        if (reports.length === 0) {
            alert("No data to export. Please run a calculation first.");
            return;
        }

        const exportService = new ExportService();
        await exportService.exportBOQ(reports);
    };

    // --- Render Helpers ---

    const renderInputStep = () => {
        return (
            <div className="max-w-4xl mx-auto mt-4">
                <BulkInputComponent onDataParsed={handlePasteSubmit} />
            </div>
        );
    };

    const toggleShopSelection = (idx) => {
        setSelectedShops(prev => {
            const newSet = new Set(prev);
            if (newSet.has(idx)) newSet.delete(idx);
            else newSet.add(idx);
            return newSet;
        });
    };

    const toggleShopExpand = (idx) => {
        setExpandedShops(prev => {
            const newSet = new Set(prev);
            if (newSet.has(idx)) newSet.delete(idx);
            else newSet.add(idx);
            return newSet;
        });
    };

    const selectAllShops = () => {
        if (selectedShops.size === reports.length) {
            setSelectedShops(new Set());
        } else {
            setSelectedShops(new Set(reports.map((_, i) => i)));
        }
    };

    const handleSelectiveExport = () => {
        const exportService = new ExportService();
        if (selectedShops.size === 0) {
            // Export all if none selected
            exportService.exportBOQ(reports);
        } else {
            // Export only selected
            const selectedReports = reports.filter((_, i) => selectedShops.has(i));
            exportService.exportBOQ(selectedReports);
        }
    };

    const renderResultStep = () => {
        const totalCostAll = reports.reduce((acc, r) => acc + (r.totalCost || 0), 0);

        return (
            <div className="max-w-6xl mx-auto animate-fade-in-up">
                {/* Header */}
                <div className="bg-white/90 backdrop-blur rounded-2xl border border-slate-200 shadow-sm mb-6 overflow-hidden">
                    <div className="bg-gradient-to-r from-energy-800 to-cyan-900 px-8 py-6 flex justify-between items-center relative overflow-hidden">
                        <div className="absolute inset-0 bg-[url('https://www.transparenttextures.com/patterns/carbon-fibre.png')] opacity-10"></div>
                        <h2 className="text-2xl font-bold text-white flex items-center gap-4 font-display relative z-10">
                            <span className="w-12 h-12 rounded-xl bg-gradient-to-br from-green-400 to-energy-600 flex items-center justify-center shadow-lg transform rotate-3 hover:rotate-0 transition-all duration-300">
                                <svg className="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M9 17V7m0 10a2 2 0 01-2 2H5a2 2 0 01-2-2V7a2 2 0 012-2h2a2 2 0 012 2m0 10a2 2 0 002 2h2a2 2 0 002-2M9 7a2 2 0 012-2h2a2 2 0 012 2m0 10V7m0 10a2 2 0 002 2h2a2 2 0 002-2V7a2 2 0 00-2-2h-2a2 2 0 00-2 2" />
                                </svg>
                            </span>
                            <div>
                                <div className="text-sm font-light text-cyan-200 uppercase tracking-widest">Engineering Report</div>
                                <div className="tracking-tight">Kết Quả Tính Toán</div>
                            </div>
                        </h2>
                        <div className="flex gap-3 relative z-10">
                            <button
                                onClick={resetWorkflow}
                                className="px-5 py-2.5 bg-white/10 hover:bg-white/20 border border-white/20 text-white rounded-xl font-medium transition-all backdrop-blur-sm flex items-center gap-2"
                            >
                                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 9v3m0 0v3m0-3h3m-3 0H9m12 0a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                                New Project
                            </button>
                            <button
                                onClick={handleSelectiveExport}
                                className="px-5 py-2.5 bg-gradient-to-r from-green-500 to-emerald-600 hover:from-green-600 hover:to-emerald-700 text-white rounded-xl font-bold shadow-lg shadow-green-900/20 flex items-center gap-2 transition-all hover:-translate-y-0.5"
                            >
                                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
                                Export Data {selectedShops.size > 0 ? `(${selectedShops.size})` : 'All'}
                            </button>
                        </div>
                    </div>

                    {/* Selection Controls */}
                    <div className="bg-slate-50/80 px-8 py-4 flex justify-between items-center border-t border-slate-200/60 backdrop-blur">
                        <div className="flex items-center gap-3">
                            <input
                                type="checkbox"
                                checked={selectedShops.size === reports.length && reports.length > 0}
                                onChange={selectAllShops}
                                className="w-5 h-5 rounded border-slate-300 text-energy-600 focus:ring-energy-500 cursor-pointer"
                            />
                            <span className="text-sm font-medium text-slate-600">
                                {selectedShops.size === 0 ? 'Chọn dự án để xuất riêng lẻ' : `Đã chọn ${selectedShops.size}/${reports.length} dự án`}
                            </span>
                        </div>
                        <button
                            onClick={() => setExpandedShops(expandedShops.size === reports.length ? new Set() : new Set(reports.map((_, i) => i)))}
                            className="text-sm text-energy-700 hover:text-energy-900 font-bold uppercase tracking-wider flex items-center gap-1 transition-colors bg-energy-50 px-3 py-1 rounded-lg border border-energy-100 hover:border-energy-300"
                        >
                            {expandedShops.size === reports.length ? (
                                <>
                                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 15l7-7 7 7" /></svg>
                                    Thu gọn
                                </>
                            ) : (
                                <>
                                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7" /></svg>
                                    Mở rộng chi tiết
                                </>
                            )}
                        </button>
                    </div>
                </div>

                {/* Shop Cards */}
                <div className="space-y-4">
                    {reports.map((report, idx) => (
                        <div key={idx} className={`bg-white rounded-xl overflow-hidden transition-all duration-300 ${selectedShops.has(idx) ? 'ring-2 ring-energy-500 shadow-md shadow-energy-100' : 'border border-gray-100 hover:border-energy-200 hover:shadow-md'}`}>
                            <div
                                className="px-6 py-4 flex items-center gap-4 cursor-pointer hover:bg-slate-50/50 transition-colors group"
                                onClick={() => toggleShopExpand(idx)}
                            >
                                <input
                                    type="checkbox"
                                    checked={selectedShops.has(idx)}
                                    onChange={(e) => { e.stopPropagation(); toggleShopSelection(idx); }}
                                    className="w-5 h-5 rounded border-gray-300 text-energy-600 focus:ring-energy-500 cursor-pointer"
                                />
                                <span className={`flex items-center justify-center w-8 h-8 rounded-full transition-colors ${expandedShops.has(idx) ? 'bg-energy-100 text-energy-700' : 'bg-slate-100 text-slate-400 group-hover:bg-white group-hover:shadow-sm'}`}>
                                    <svg className={`w-5 h-5 transform transition-transform ${expandedShops.has(idx) ? 'rotate-180' : ''}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7" />
                                    </svg>
                                </span>

                                <div className="flex-1">
                                    <h4 className="font-bold text-slate-800 text-lg flex items-center gap-2">
                                        {report.projectName || report.projectId}
                                        {report.config.pvModel && <span className="text-[10px] font-mono font-normal bg-slate-100 px-2 py-0.5 rounded text-slate-500 border border-slate-200">{report.config.pvModel}</span>}
                                    </h4>
                                    <div className="text-sm text-slate-500 flex items-center gap-4 mt-1">
                                        <div className="flex items-center gap-1.5">
                                            <span className="w-2 h-2 rounded-full bg-energy-500"></span>
                                            <span className="text-energy-700 font-mono font-bold">{report.config.dcPower} kWp</span>
                                        </div>
                                        <span className="text-slate-300">|</span>
                                        <div className="flex items-center gap-1.5">
                                            <svg className="w-4 h-4 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10" /></svg>
                                            <span className="font-mono">{report.items.length} Items</span>
                                        </div>
                                    </div>
                                </div>
                                <div className="text-right">
                                    <div className="text-xs text-slate-400 font-medium mb-1">DỰ TOÁN (ƯỚC TÍNH)</div>
                                    <span className="font-mono font-bold text-xl text-slate-900 bg-slate-50 px-3 py-1 rounded-lg border border-slate-100 block">
                                        {new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(report.totalCost)}
                                    </span>
                                </div>
                            </div>

                            {/* Collapsible Table */}
                            {expandedShops.has(idx) && (
                                <div className="max-h-96 overflow-y-auto border-t border-slate-100 custom-scrollbar bg-slate-50/30">
                                    <table className="w-full text-sm">
                                        <thead className="bg-slate-50 sticky top-0 z-10 shadow-sm">
                                            <tr className="text-slate-500 text-xs font-bold uppercase tracking-wider text-left">
                                                <th className="py-2.5 px-6 w-24">Group</th>
                                                <th className="py-2.5 px-6">Hạng Mục / Thiết Bị</th>
                                                <th className="py-2.5 px-6 text-right w-32">Số Lượng</th>
                                                <th className="py-2.5 px-6 text-center w-20">ĐVT</th>
                                            </tr>
                                        </thead>
                                        <tbody className="divide-y divide-slate-100">
                                            {report.items.map((item, i) => (
                                                <tr key={i} className="hover:bg-energy-50/40 transition-colors">
                                                    <td className="py-2.5 px-6 text-[10px] font-bold text-slate-400 uppercase tracking-widest">{item.group}</td>
                                                    <td className="py-2.5 px-6 font-medium text-slate-700 text-sm">{item.name}</td>
                                                    <td className="py-2.5 px-6 text-right font-mono font-bold text-energy-700">{new Intl.NumberFormat('vi-VN').format(item.quantity)}</td>
                                                    <td className="py-2.5 px-6 text-center text-xs text-slate-500 font-medium">{item.unit}</td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                            )}
                        </div>
                    ))}
                </div>
            </div>
        );
    }

    return (
        <div className="w-full pb-8 animate-fade-in">


            <div className="container mx-auto px-4">
                {error && (
                    <div className="mb-6 mx-auto max-w-4xl bg-red-50 border border-red-200 text-red-700 px-6 py-4 rounded-xl shadow-sm flex items-center gap-3 animate-fade-in-up">
                        <svg className="w-6 h-6 text-red-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                        <span className="font-bold">Error:</span> {error}
                        <button onClick={() => setError(null)} className="ml-auto text-sm font-bold hover:underline">Dismiss</button>
                    </div>
                )}

                {activeStep === 0 && renderInputStep()}
                {activeStep === 1 && renderResultStep()}
            </div>
        </div>
    );


}
