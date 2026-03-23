import React, { useState, useEffect, useMemo } from 'react';
import { ArrowRight, ArrowLeft, History, AlertCircle } from 'lucide-react';

const LOAN_TYPE_SUGGESTIONS = ['Term Loan', 'Cash Credit', 'Overdraft', 'Gold Loan', 'Housing Loan', 'Vehicle Loan', 'Education Loan', 'MSME Loan'];

export default function MissingMappingsForm({ missingCodes, existingMappings = {}, onComplete, onBack }) {
    const [manualMappings, setManualMappings] = useState(() => {
        const initial = {};
        missingCodes.forEach(code => {
            initial[code] = existingMappings[code] || '';
        });
        return initial;
    });

    // Sort codes: unmapped (empty) at the top, then alphabetically
    const sortedCodes = useMemo(() => {
        return [...missingCodes].sort((a, b) => {
            const aVal = manualMappings[a] || '';
            const bVal = manualMappings[b] || '';

            // If one is empty and other isn't, empty comes first
            if (!aVal && bVal) return -1;
            if (aVal && !bVal) return 1;

            // Otherwise alphabetical
            return a.localeCompare(b);
        });
    }, [missingCodes, manualMappings]);

    const handleChange = (code, val) => setManualMappings(prev => ({ ...prev, [code]: val }));

    const handleComplete = () => {
        const resolved = {};
        Object.entries(manualMappings).forEach(([code, type]) => {
            if (type.trim()) resolved[code] = type.trim();
        });
        onComplete(resolved);
    };

    const filledCount = Object.values(manualMappings).filter(v => v.trim()).length;

    return (
        <div className="card card-body animate-slide" style={{ maxWidth: 660, margin: '0 auto' }}>
            {/* Header */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.75rem' }}>
                <div style={{
                    width: 44, height: 44,
                    background: 'var(--p-sky)',
                    borderRadius: 11,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    flexShrink: 0
                }}>
                    <History size={20} color="#0369A1" />
                </div>
                <div>
                    <h2 className="page-title" style={{ fontSize: '1.375rem' }}>Review Missing Mappings</h2>
                    <p style={{ fontSize: '0.8375rem', color: 'var(--text-muted)', marginTop: 2 }}>
                        These product codes were not found in your Excel mapping file.
                    </p>
                </div>
            </div>

            <hr className="divider" />

            <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '0.875rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <AlertCircle size={14} />
                    <span>Assign loan types or review history matches below.</span>
                </div>
                <span className="badge badge-sky">{filledCount} / {missingCodes.length} mapped</span>
            </div>

            {/* Codes list */}
            <div style={{ maxHeight: 400, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '0.625rem', marginBottom: '1.5rem', paddingRight: '0.25rem' }}>
                {sortedCodes.map((code, idx) => {
                    const isPreFilled = !!existingMappings[code];
                    const currentVal = manualMappings[code] || '';

                    return (
                        <div key={code} style={{
                            display: 'grid',
                            gridTemplateColumns: 'minmax(120px, 1fr) 2fr',
                            alignItems: 'center',
                            gap: '1rem',
                            padding: '12px 16px',
                            background: isPreFilled ? 'rgba(224, 242, 254, 0.3)' : 'var(--surface-hover)',
                            borderRadius: 'var(--r-sm)',
                            border: isPreFilled ? '1px solid var(--p-sky-d)' : '1px solid var(--border)',
                            transition: 'all 0.2s ease'
                        }}>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                                <span style={{
                                    fontFamily: "'Inter', monospace",
                                    fontWeight: 700,
                                    fontSize: '0.875rem',
                                    color: 'var(--text-heading)',
                                    background: isPreFilled ? 'var(--p-sky)' : 'var(--p-lavender)',
                                    padding: '2px 8px',
                                    borderRadius: 'var(--r-pill)',
                                    width: 'fit-content'
                                }}>{code}</span>
                                {isPreFilled && (
                                    <span style={{ fontSize: '0.65rem', color: '#0369A1', fontWeight: 600, textTransform: 'uppercase' }}>
                                        History Match
                                    </span>
                                )}
                            </div>
                            <div style={{ position: 'relative' }}>
                                <input
                                    list={`suggestions-${idx}`}
                                    className="form-input"
                                    style={{
                                        fontSize: '0.875rem',
                                        padding: '8px 12px',
                                        borderColor: isPreFilled && currentVal === existingMappings[code] ? 'var(--p-sky-d)' : undefined
                                    }}
                                    placeholder="Select loan type…"
                                    value={currentVal}
                                    onChange={e => handleChange(code, e.target.value)}
                                />
                                <datalist id={`suggestions-${idx}`}>
                                    {LOAN_TYPE_SUGGESTIONS.map(s => <option key={s} value={s} />)}
                                </datalist>
                            </div>
                        </div>
                    );
                })}
            </div>

            <div style={{ display: 'flex', gap: '0.75rem', justifyContent: 'flex-end', paddingTop: '0.5rem' }}>
                <button className="btn btn-secondary" onClick={onBack}><ArrowLeft size={15} /> Back</button>
                <button className="btn btn-primary" onClick={handleComplete}>
                    Apply Mappings <ArrowRight size={15} />
                </button>
            </div>
        </div>
    );
}
