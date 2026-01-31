import { useState } from 'react';
import CalculatorPage from './presentation/pages/CalculatorPage';
import cpsLogo from './assets/cps_logo.png';

function App() {
  const [resetKey, setResetKey] = useState(0);

  const handleReset = () => {
    setResetKey(prev => prev + 1);
  };

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 font-sans text-slate-900 selection:bg-energy-200 selection:text-energy-900">
      {/* Header with CPS Branding */}
      <header className="bg-white/80 backdrop-blur-md shadow-sm border-b border-slate-200 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 py-4">
          <div className="flex flex-row items-center justify-center gap-5">
            {/* CPS Logo - Clickable Reset */}
            <button
              onClick={handleReset}
              className="cursor-pointer focus:outline-none transition-transform active:scale-95 group"
              title="Quay lại Dashboard"
            >
              <img
                src={cpsLogo}
                alt="CPS Logo"
                className="h-24 w-auto object-contain transition-all duration-300 drop-shadow-sm group-hover:drop-shadow-md group-hover:scale-105"
              />
            </button>

            {/* Title Block with Glassmorphism & Energy Gradient */}
            <div className="relative group p-4 rounded-2xl overflow-hidden transition-all duration-500 border border-transparent hover:border-energy-100 hover:shadow-lg hover:shadow-energy-100/50">
              {/* Animated Background */}
              <div className="absolute inset-0 bg-gradient-to-br from-energy-50/50 via-white to-cyan-50/50 opacity-0 group-hover:opacity-100 transition-opacity duration-500"></div>

              <div className="relative border-l-4 border-energy-500 pl-6">
                <h1 className="text-4xl font-extrabold tracking-tight uppercase bg-gradient-to-r from-blue-700 via-emerald-400 to-blue-700 bg-clip-text text-transparent animate-gradient-x font-display pb-1">
                  Solar BOQ & BOM Engine
                </h1>
                <p className="text-slate-500 text-lg mt-2 font-medium tracking-wide flex items-center gap-2 font-mono text-sm">
                  Bill of Quantities & Materials Calculator
                  <span className="relative flex h-2.5 w-2.5">
                    <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-energy-400 opacity-75"></span>
                    <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-energy-500"></span>
                  </span>
                </p>
              </div>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 w-full max-w-7xl mx-auto px-6 py-2 overflow-hidden">
        <CalculatorPage key={resetKey} />
      </main>

      {/* Footer */}
      <footer className="bg-slate-900 border-t border-slate-800 mt-auto text-slate-400 py-8">
        <div className="max-w-7xl mx-auto px-6 flex flex-col md:flex-row justify-between items-center gap-4">
          <div className="text-sm">
            <span className="font-bold text-slate-200">CPS Solar Solutions</span> © {new Date().getFullYear()}
            <span className="mx-2">•</span>
            <span className="text-xs font-mono">Engineering Division</span>
          </div>

          <div className="flex items-center gap-3 bg-slate-800/50 px-4 py-2 rounded-full border border-slate-700">
            <div className="flex items-center gap-1.5">
              <span className="w-2 h-2 rounded-full bg-energy-500 shadow-[0_0_8px_rgba(20,184,166,0.6)]"></span>
              <span className="text-energy-400 text-xs font-mono font-bold tracking-wider">SYSTEM READY</span>
            </div>
            <span className="text-slate-600 text-xs">|</span>
            <span className="text-xs text-slate-500 font-mono">v2.1.0-2026</span>
          </div>
        </div>
      </footer>
    </div>
  );
}

export default App;
