import React, { useState } from 'react';
import { ArrowRight, ArrowLeft, Tag, Info } from 'lucide-react';

export default function FacilityColumnSelector({ availableColumns, onSelect, onBack }) {
    const [selected, setSelected] = useState('');

    return (
        <div className="card card-body animate-slide" style={{ maxWidth: 540, margin: '0 auto' }}>
            {/* Header */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.75rem' }}>
                <div style={{
                    width: 44, height: 44,
                    background: 'var(--p-peach)',
                    borderRadius: 11,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    flexShrink: 0
                }}>
                    <Tag size={20} color="#92400E" />
                </div>
                <div>
                    <h2 className="page-title" style={{ fontSize: '1.375rem' }}>Product Mapping</h2>
                    <p style={{ fontSize: '0.8375rem', color: 'var(--text-muted)', marginTop: 2 }}>
                        Identify which column holds your product / facility codes.
                    </p>
                </div>
            </div>

            <hr className="divider" />

            <div className="form-group" style={{ marginBottom: '1.25rem' }}>
                <label className="form-label">Select Product Code Column</label>
                <select
                    className="form-input"
                    value={selected}
                    onChange={e => setSelected(e.target.value)}
                >
                    <option value="">— Select a column —</option>
                    {availableColumns.map(col => (
                        <option key={col} value={col}>{col}</option>
                    ))}
                </select>
            </div>

            <div className="alert alert-info" style={{ marginBottom: '1.5rem' }}>
                <Info size={14} style={{ flexShrink: 0, marginTop: 1 }} />
                <span>This column's values will be matched against your Product Mapping file to assign loan types (TL, CC, OD, Gold, etc.) to each account.</span>
            </div>

            <div style={{ display: 'flex', gap: '0.75rem', justifyContent: 'flex-end' }}>
                <button className="btn btn-secondary" onClick={onBack}><ArrowLeft size={15} /> Back</button>
                <button className="btn btn-primary" onClick={() => selected && onSelect(selected)} disabled={!selected}>
                    Confirm Mapping <ArrowRight size={15} />
                </button>
            </div>
        </div>
    );
}
