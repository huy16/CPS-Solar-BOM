
import React, { useState } from 'react';
import { ProjectConfig } from '../../domain/entities/ProjectConfig';
import { DI } from '../../infrastructure/di/Container';
import DatalinkImporter from '../components/DatalinkImporter';
import BulkInputComponent from '../components/BulkInputComponent';
import WorkflowStepper from '../components/WorkflowStepper';
import equipmentData from '../../data/equipment_data.json';
import { ExportService } from '../../services/ExportService';

// Simple Error Boundary for Debugging
class ErrorBoundary extends React.Component {
    constructor(props) {
        super(props);
        this.state = { hasError: false, error: null };
    }

    static getDerivedStateFromError(error) {
        return { hasError: true, error };
    }

    componentDidCatch(error, errorInfo) {
        console.error("ErrorBoundary caught an error", error, errorInfo);
    }

    render() {
        if (this.state.hasError) {
            return (
                <div className="p-4 bg-red-100 border border-red-400 text-red-700 rounded-lg m-4">
                    <h2 className="font-bold text-lg mb-2">Something went wrong.</h2>
                    <details className="whitespace-pre-wrap font-mono text-sm">
                        {this.state.error && this.state.error.toString()}
                    </details>
                </div>
            );
        }
        return this.props.children;
    }
}

// Wrap the export
export default function CalculatorPageWrapper() {
    return (
        <ErrorBoundary>
            <CalculatorPage />
        </ErrorBoundary>
    );
}

function CalculatorPage() {
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
    const [viewDetailReport, setViewDetailReport] = useState(null);

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
        const totalCostAll = (reports || []).filter(r => r).reduce((acc, r) => acc + (r?.totalCost || 0), 0);

        return (
            <div className="max-w-6xl mx-auto animate-fade-in-up">
                {/* Header */}
                <div className="bg-white/90 backdrop-blur rounded-2xl border border-slate-200 shadow-sm mb-4 overflow-hidden">
                    <div className="bg-gradient-to-r from-energy-800 to-cyan-900 px-4 md:px-6 py-3 md:py-4 flex flex-col md:flex-row justify-between items-center gap-3 relative overflow-hidden">
                        <div className="absolute inset-0 bg-[url('https://www.transparenttextures.com/patterns/carbon-fibre.png')] opacity-10"></div>
                        <h2 className="text-lg md:text-xl font-bold text-white flex items-center gap-3 md:gap-4 font-display relative z-10">
                            <span className="w-8 h-8 md:w-10 md:h-10 rounded-lg bg-gradient-to-br from-green-400 to-energy-600 flex items-center justify-center shadow-lg transform rotate-3 hover:rotate-0 transition-all duration-300">
                                <svg className="w-4 h-4 md:w-5 md:h-5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M9 17V7m0 10a2 2 0 01-2 2H5a2 2 0 01-2-2V7a2 2 0 012-2h2a2 2 0 012 2m0 10a2 2 0 002 2h2a2 2 0 002-2M9 7a2 2 0 012-2h2a2 2 0 012 2m0 10V7m0 10a2 2 0 002 2h2a2 2 0 002-2V7a2 2 0 00-2-2h-2a2 2 0 00-2 2" />
                                </svg>
                            </span>
                            <div>
                                <div className="text-[9px] font-light text-cyan-200 uppercase tracking-widest leading-none">Engineering Report</div>
                                <div className="tracking-tight mt-0.5">Kết Quả Tính Toán</div>
                            </div>
                        </h2>
                        <div className="flex flex-col sm:flex-row gap-2 w-full md:w-auto relative z-10">
                            <button
                                onClick={resetWorkflow}
                                className="px-4 py-2 bg-white/10 hover:bg-white/20 border border-white/20 text-white rounded-lg font-medium transition-all backdrop-blur-sm flex items-center justify-center md:justify-start gap-2 text-xs md:text-sm"
                            >
                                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 9v3m0 0v3m0-3h3m-3 0H9m12 0a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                                New Project
                            </button>
                            <button
                                onClick={handleSelectiveExport}
                                className="px-4 py-2 bg-gradient-to-r from-green-500 to-emerald-600 hover:from-green-600 hover:to-emerald-700 text-white rounded-lg font-bold shadow-lg shadow-green-900/20 flex items-center justify-center md:justify-start gap-2 transition-all hover:-translate-y-0.5 text-xs md:text-sm"
                            >
                                <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
                                Export Data {selectedShops.size > 0 ? `(${selectedShops.size})` : 'All'}
                            </button>
                        </div>
                    </div>

                    {/* Selection Controls */}
                    <div className="bg-slate-50/80 px-4 md:px-6 py-2 flex justify-between items-center border-t border-slate-200/60 backdrop-blur">
                        <div className="flex items-center gap-2">
                            <input
                                type="checkbox"
                                checked={selectedShops.size === reports.length && reports.length > 0}
                                onChange={selectAllShops}
                                className="w-4 h-4 rounded border-slate-300 text-energy-600 focus:ring-energy-500 cursor-pointer"
                            />
                            <span className="text-[10px] md:text-xs font-medium text-slate-600">
                                {selectedShops.size === 0 ? 'Chọn dự án để xuất riêng lẻ' : `Đã chọn ${selectedShops.size}/${reports.length} dự án`}
                            </span>
                        </div>
                    </div>
                </div>

                {/* Grid View */}
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 xl:grid-cols-5 gap-3">
                    {reports.map((report, idx) => (
                        <div
                            key={idx}
                            className={`bg-white rounded-xl p-3 border transition-all duration-300 group cursor-pointer relative hover:-translate-y-1 hover:shadow-lg ${selectedShops.has(idx) ? 'border-energy-500 ring-2 ring-energy-500/20 shadow-md' : 'border-slate-200 hover:border-energy-300'}`}
                            onClick={() => setViewDetailReport(report)}
                        >
                            <div className="flex justify-between items-start mb-2">
                                <div className="bg-slate-100 rounded px-1.5 py-0.5 text-[10px] font-mono font-bold text-slate-500 border border-slate-200 truncate max-w-[70%]">
                                    {report.projectId}
                                </div>
                                <div
                                    onClick={(e) => { e.stopPropagation(); toggleShopSelection(idx); }}
                                    className={`w-5 h-5 rounded-full border flex items-center justify-center transition-colors ${selectedShops.has(idx) ? 'bg-energy-500 border-energy-500 text-white' : 'bg-white border-slate-300 text-transparent hover:border-energy-400'}`}
                                >
                                    <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M5 13l4 4L19 7" /></svg>
                                </div>
                            </div>

                            <h4 className="font-bold text-slate-800 text-sm mb-1 truncate" title={report?.projectName}>
                                {report?.projectName || report?.projectId}
                            </h4>

                            <div className="flex items-center gap-2 mb-3">
                                <span className="text-[10px] font-bold text-white bg-gradient-to-r from-blue-500 to-cyan-500 px-2 py-0.5 rounded-full shadow-sm shadow-blue-200">
                                    {report?.config?.dcPower} kWp
                                </span>
                                {report?.config?.pvModel && (
                                    <span className="text-[9px] text-slate-400 font-medium truncate flex-1 text-right" title={report.config.pvModel}>
                                        {String(report.config.pvModel).split('-').pop()}
                                    </span>
                                )}
                            </div>

                            <div className="pt-2 border-t border-slate-100 flex justify-between items-center">
                                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">{report?.items?.length || 0} Items</span>
                                <span className="w-6 h-6 rounded-full bg-slate-50 text-slate-400 group-hover:bg-energy-50 group-hover:text-energy-600 flex items-center justify-center transition-colors">
                                    <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" /><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" /></svg>
                                </span>
                            </div>
                        </div>
                    ))}
                </div>

                {/* Detail Modal */}
                {viewDetailReport && (
                    <div className="fixed inset-0 z-[9999] flex items-center justify-center p-2 bg-slate-900/60 backdrop-blur-sm animate-fade-in">
                        <div className="bg-white rounded-2xl md:rounded-3xl shadow-2xl w-full max-w-5xl max-h-[95vh] flex flex-col overflow-hidden animate-zoom-in">
                            <div className="px-5 py-3 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                                <div className="flex flex-col md:flex-row md:items-center gap-1 md:gap-4">
                                    <div className="flex items-center gap-2">
                                        <h3 className="text-base md:text-lg font-bold text-slate-800 font-display">{viewDetailReport?.projectName}</h3>
                                        <span className="bg-energy-100 text-energy-700 text-[9px] font-bold px-1.5 py-0.5 rounded border border-energy-200">{viewDetailReport?.projectId}</span>
                                    </div>
                                    <div className="flex items-center gap-3 text-[11px] md:text-xs text-slate-500 font-medium">
                                        <span className="flex items-center gap-1"><strong className="text-slate-700">{viewDetailReport?.config?.dcPower}</strong> kWp</span>
                                        <span className="flex items-center gap-1"><strong className="text-slate-700">{viewDetailReport?.config?.panelCount}</strong> Panels</span>
                                        <span className="flex items-center gap-1 bg-slate-100 px-1.5 rounded text-[10px]">{viewDetailReport?.config?.pvModel}</span>
                                    </div>
                                </div>
                                <button ref={el => el?.focus()} onClick={() => setViewDetailReport(null)} className="w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-500 hover:text-red-500 flex items-center justify-center transition-all">
                                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
                                </button>
                            </div>

                            <div className="flex-1 overflow-y-auto custom-scrollbar p-3 md:p-4 bg-slate-50/30">
                                <table className="w-full text-xs md:text-sm rounded-lg overflow-hidden ring-1 ring-slate-200">
                                    <thead className="bg-slate-100 sticky top-0 shadow-sm">
                                        <tr className="text-slate-600 text-[9px] font-bold uppercase tracking-wider text-left">
                                            <th className="py-2 px-4 w-24 md:w-32">Group</th>
                                            <th className="py-2 px-4">Hạng Mục / Thiết Bị</th>
                                            <th className="py-2 px-4 text-right w-24 md:w-32">Số Lượng</th>
                                            <th className="py-2 px-4 text-center w-20 md:w-24">ĐVT</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y divide-slate-100 bg-white">
                                        {viewDetailReport?.items?.map((item, i) => (
                                            <tr key={i} className="hover:bg-blue-50/30 transition-colors">
                                                <td className="py-2 px-4 text-[9px] font-bold text-slate-400 uppercase tracking-widest">{item.group}</td>
                                                <td className="py-2 px-4 font-medium text-slate-700 text-[11px] md:text-sm">{item.name}</td>
                                                <td className="py-2 px-4 text-right font-mono font-bold text-blue-600 text-[11px] md:text-sm">{new Intl.NumberFormat('vi-VN').format(item.quantity)}</td>
                                                <td className="py-2 px-4 text-center text-[10px] md:text-xs text-slate-500">{item.unit}</td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>

                            <div className="px-5 py-2.5 border-t border-slate-100 bg-white flex justify-end gap-3">
                                <button onClick={() => setViewDetailReport(null)} className="px-6 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl transition-all text-sm">Đóng</button>
                            </div>
                        </div>
                    </div>
                )}
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
