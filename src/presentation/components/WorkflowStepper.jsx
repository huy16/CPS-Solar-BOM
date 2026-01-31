import React from 'react';

export default function WorkflowStepper({ currentStep, steps }) {
    return (
        <div className="w-full py-8">
            <div className="flex items-center justify-center relative z-10">
                {steps.map((step, index) => {
                    const isCompleted = index < currentStep;
                    const isCurrent = index === currentStep;

                    return (
                        <div key={index} className="flex items-center">
                            {/* Step Circle */}
                            <div className="relative flex flex-col items-center group">
                                <div
                                    className={`w-12 h-12 rounded-2xl flex items-center justify-center border-2 transition-all duration-500 shadow-sm
                                        ${isCompleted ? 'bg-energy-500 border-energy-500 text-white shadow-energy-200' :
                                            isCurrent ? 'bg-white border-cyan-500 text-cyan-600 ring-4 ring-cyan-50 shadow-lg scale-110' :
                                                'bg-white border-slate-200 text-slate-300'}
                                    `}
                                >
                                    {isCompleted ? (
                                        <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M5 13l4 4L19 7" />
                                        </svg>
                                    ) : (
                                        <span className={`font-bold text-lg font-mono ${isCurrent ? 'animate-pulse' : ''}`}>{index + 1}</span>
                                    )}
                                </div>
                                <div className={`absolute top-14 whitespace-nowrap text-sm font-bold tracking-wide transition-colors duration-300 font-display
                                    ${isCompleted ? 'text-energy-700' : isCurrent ? 'text-cyan-700' : 'text-slate-400'}`}>
                                    {step.label}
                                </div>
                            </div>

                            {/* Connector Line */}
                            {index < steps.length - 1 && (
                                <div className={`w-32 h-1 mx-4 rounded-full transition-all duration-500
                                    ${index < currentStep ? 'bg-gradient-to-r from-energy-500 to-cyan-500' : 'bg-slate-100'}
                                `}></div>
                            )}
                        </div>
                    );
                })}
            </div>
        </div>
    );
}
