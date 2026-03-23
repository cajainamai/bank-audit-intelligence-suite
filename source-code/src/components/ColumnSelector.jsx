import React, { useState, useEffect } from 'react';
import { ArrowRight, ArrowLeft, Search, CheckSquare, Square } from 'lucide-react';

const ACCENT_COLORS = [
    'var(--p-lavender)', 'var(--p-sky)', 'var(--p-mint)', 'var(--p-rose)',
    'var(--p-peach)', 'var(--p-lilac)'
];

export default function ColumnSelector({ headers, onProceed, onBack }) {
    const [selectedColumns, setSelectedColumns] = useState(() => {
        const saved = localStorage.getItem('dataSummarizer_selectedColumns');
        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                return headers.reduce((acc, h) => ({ ...acc, [h]: parsed[h] !== undefined ? parsed[h] : true }), {});
            } catch (e) { /* ignore */ }
        }
        return headers.reduce((acc, h) => ({ ...acc, [h]: true }), {});
    });
    const [search, setSearch] = useState('');

    useEffect(() => {
        localStorage.setItem('dataSummarizer_selectedColumns', JSON.stringify(selectedColumns));
    }, [selectedColumns]);

    const toggle = (h) => setSelectedColumns(prev => ({ ...prev, [h]: !prev[h] }));
    const selectAll = () => setSelectedColumns(headers.reduce((acc, h) => ({ ...acc, [h]: true }), {}));
    const deselectAll = () => setSelectedColumns(headers.reduce((acc, h) => ({ ...acc, [h]: false }), {}));

    const filtered = headers.filter(h => h.toLowerCase().includes(search.toLowerCase()));
    const selectedCount = Object.values(selectedColumns).filter(Boolean).length;

    const handleProceed = () => {
        const cols = headers.filter(h => selectedColumns[h]);
        onProceed(cols);
    };

    return (
        <div className="card card-body animate-slide" style={{ maxWidth: 700, margin: '0 auto' }}>
            {/* Header */}
            <div style={{ marginBottom: '1.75rem' }}>
                <h2 className="page-title" style={{ fontSize: '1.375rem' }}>Select Columns</h2>
                <p className="page-subtitle" style={{ marginTop: 4 }}>
                    Choose which columns from your data file to carry into the report.
                </p>
            </div>

            {/* Toolbar */}
            <div style={{ display: 'flex', gap: '0.75rem', alignItems: 'center', marginBottom: '1.25rem', flexWrap: 'wrap' }}>
                <div style={{ position: 'relative', flex: 1, minWidth: 180 }}>
                    <Search size={14} color="var(--text-faint)" style={{ position: 'absolute', left: 11, top: '50%', transform: 'translateY(-50%)' }} />
                    <input
                        className="form-input"
                        placeholder="Search columns…"
                        value={search}
                        onChange={e => setSearch(e.target.value)}
                        style={{ paddingLeft: 32 }}
                    />
                </div>
                <button className="btn btn-secondary btn-sm" onClick={selectAll}>Select All</button>
                <button className="btn btn-secondary btn-sm" onClick={deselectAll}>Deselect All</button>
                <span style={{ fontSize: '0.8125rem', color: 'var(--text-muted)', fontWeight: 500, marginLeft: 'auto', whiteSpace: 'nowrap' }}>
                    {selectedCount} / {headers.length} selected
                </span>
            </div>

            {/* Column grid */}
            <div style={{
                display: 'flex',
                flexWrap: 'wrap',
                gap: '0.5rem',
                maxHeight: 340,
                overflowY: 'auto',
                padding: '0.25rem 0',
                marginBottom: '1.5rem'
            }}>
                {filtered.map((h, i) => {
                    const isSelected = selectedColumns[h];
                    return (
                        <div
                            key={h}
                            onClick={() => toggle(h)}
                            style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: 6,
                                padding: '5px 11px',
                                borderRadius: 'var(--r-pill)',
                                border: '1.5px solid',
                                fontSize: '0.8125rem',
                                fontWeight: 500,
                                cursor: 'pointer',
                                transition: 'var(--t-fast)',
                                userSelect: 'none',
                                background: isSelected ? ACCENT_COLORS[i % ACCENT_COLORS.length] : 'var(--surface-hover)',
                                borderColor: isSelected ? 'transparent' : 'var(--border-strong)',
                                color: isSelected ? 'var(--text-heading)' : 'var(--text-muted)',
                            }}
                        >
                            {isSelected
                                ? <CheckSquare size={13} color="var(--brand-primary)" />
                                : <Square size={13} color="var(--text-faint)" />
                            }
                            {h}
                        </div>
                    );
                })}
                {filtered.length === 0 && (
                    <p style={{ color: 'var(--text-faint)', fontSize: '0.875rem' }}>No columns match your search.</p>
                )}
            </div>

            {/* Actions */}
            <div style={{ display: 'flex', gap: '0.75rem', justifyContent: 'flex-end' }}>
                <button className="btn btn-secondary" onClick={onBack}><ArrowLeft size={15} /> Back</button>
                <button className="btn btn-primary" onClick={handleProceed} disabled={selectedCount === 0}>
                    Continue <ArrowRight size={15} />
                </button>
            </div>
        </div>
    );
}
