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
        <div className="max-w-7xl mx-auto px-4 md:px-6 py-2 md:py-2">
          <div className="flex flex-col md:flex-row items-center justify-center gap-1 md:gap-0">
            {/* CPS Logo - Clickable Reset */}
            <button
              onClick={handleReset}
              className="cursor-pointer focus:outline-none transition-transform active:scale-95 group"
              title="Quay lại Dashboard"
            >
              <img
                src={cpsLogo}
                alt="CPS Logo"
                className="h-16 md:h-20 lg:h-24 w-auto object-contain transition-all duration-300 drop-shadow-sm group-hover:drop-shadow-md group-hover:scale-105"
              />
            </button>

            {/* Vertical Divider - Hidden on Mobile */}
            <div className="hidden md:block h-12 w-px bg-slate-200 mx-1"></div>

            {/* Title Block with Glassmorphism & Energy Gradient */}
            <div className="relative group p-1 md:p-2 rounded-2xl overflow-hidden transition-all duration-500 border border-transparent hover:border-energy-100 hover:shadow-lg hover:shadow-energy-100/50 text-center md:text-left">
              {/* Animated Background */}
              <div className="absolute inset-0 bg-gradient-to-br from-energy-50/50 via-white to-cyan-50/50 opacity-0 group-hover:opacity-100 transition-opacity duration-500"></div>

              <div className="relative pl-0 flex flex-col items-center md:items-start">
                <div className="flex flex-wrap items-center justify-center md:justify-start gap-1 md:gap-2">
                  <h1 className="text-lg sm:text-xl md:text-2xl lg:text-3xl font-extrabold tracking-tight uppercase bg-gradient-to-r from-blue-700 via-emerald-400 to-blue-700 bg-clip-text text-transparent animate-gradient-x font-display pb-0.5 pr-2">
                    Solar BOQ & BOM Engine
                  </h1>
                  {/* Lightning Bolt Icon */}
                  <svg className="w-4 h-4 md:w-6 md:h-6 animate-pulse drop-shadow-[0_0_12px_rgba(239,68,68,0.4)] shrink-0" viewBox="0 0 20 20">
                    <defs>
                      <linearGradient id="boltGradient" x1="0%" y1="0%" x2="0%" y2="100%">
                        <stop offset="0%" stopColor="#ef4444" />
                        <stop offset="100%" stopColor="#fbbf24" />
                      </linearGradient>
                    </defs>
                    <path fill="url(#boltGradient)" fillRule="evenodd" d="M11.3 1.046A1 1 0 0112 2v5h4a1 1 0 01.82 1.573l-7 10A1 1 0 018 18v-5H4a1 1 0 01-.82-1.573l7-10a1 1 0 011.12-.38z" clipRule="evenodd" />
                  </svg>
                </div>
                <p className="text-slate-500 text-[10px] sm:text-xs md:text-sm font-medium tracking-wide flex items-center gap-2 font-mono">
                  Bill of Quantities & Materials Calculator
                  <span className="relative flex h-1.5 w-1.5 md:h-2 md:w-2">
                    <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-energy-400 opacity-75"></span>
                    <span className="relative inline-flex rounded-full h-1.5 w-1.5 md:h-2 md:w-2 bg-energy-500"></span>
                  </span>
                </p>
              </div>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="flex-1 w-full max-w-7xl mx-auto px-4 md:px-6 py-1 overflow-hidden">
        <CalculatorPage key={resetKey} />
      </main>

      {/* Footer */}
      <footer className="bg-slate-900 border-t border-slate-800 mt-auto text-slate-400 py-3 md:py-4">
        <div className="max-w-7xl mx-auto px-6 flex flex-col md:flex-row justify-between items-center gap-2 md:gap-4">
          <div className="text-[10px] md:text-xs">
            <span className="font-bold text-slate-200">CPS Solar Solutions</span> © {new Date().getFullYear()}
            <span className="mx-2">•</span>
            <span className="text-[10px] font-mono opacity-60">Engineering Division</span>
          </div>

          <div className="flex items-center gap-2 md:gap-3 bg-slate-800/50 px-3 md:px-4 py-1 md:py-1.5 rounded-full border border-slate-700">
            <div className="flex items-center gap-1.5">
              <span className="w-1.5 h-1.5 rounded-full bg-energy-500 shadow-[0_0_8px_rgba(20,184,166,0.6)]"></span>
              <span className="text-energy-400 text-[9px] md:text-[10px] font-mono font-bold tracking-wider">SYSTEM READY</span>
            </div>
            <span className="text-slate-600 text-xs">|</span>
            <span className="text-[9px] md:text-[10px] text-slate-500 font-mono">v2.1.0-2026</span>
          </div>
        </div>
      </footer>
    </div>
  );
}

export default App;
