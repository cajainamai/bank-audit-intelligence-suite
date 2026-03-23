import React, { useState } from 'react';
import { Calendar, ArrowRight, Info } from 'lucide-react';

export default function AuditPeriodSelector({ onProceed }) {
    const [fromDate, setFromDate] = useState('');
    const [toDate, setToDate] = useState('');
    const [npaAvailable, setNpaAvailable] = useState('no'); // 'yes' | 'no'
    const [npaSource, setNpaSource] = useState(''); // 'same' | 'separate'

    const canProceed = fromDate && toDate && new Date(toDate) > new Date(fromDate);

    const formatDate = (d) => d ? new Date(d).toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' }) : '';

    const handleFYChange = (e) => {
        const fy = e.target.value;
        if (fy) {
            const startYear = parseInt(fy.substring(0, 4), 10);
            setFromDate(`${startYear}-04-01`);
            setToDate(`${startYear + 1}-03-31`);
        }
    };

    const fyOptions = [
        "2023-24", "2024-25", "2025-26", "2026-27",
        "2027-28", "2028-29", "2029-30"
    ];

    return (
        <div className="card card-body animate-slide" style={{ maxWidth: 540, margin: '0 auto' }}>
            {/* Header */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.75rem' }}>
                <div style={{
                    width: 44, height: 44,
                    background: 'var(--p-lavender)',
                    borderRadius: 11,
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    flexShrink: 0
                }}>
                    <Calendar size={22} color="#4C1D95" />
                </div>
                <div>
                    <h2 className="page-title" style={{ fontSize: '1.375rem' }}>Set Audit Period</h2>
                    <p style={{ fontSize: '0.8375rem', color: 'var(--text-muted)', marginTop: 2 }}>
                        Define the From and To dates for this audit cycle.
                    </p>
                </div>
            </div>

            <hr className="divider" />

            {/* FY Selection (Optional) */}
            <div className="form-group" style={{ marginBottom: '1.25rem' }}>
                <label className="form-label">Financial Year <span style={{ color: 'var(--text-faint)', fontWeight: 400 }}>(Optional - Auto-fills dates)</span></label>
                <select className="form-input" onChange={handleFYChange} defaultValue="">
                    <option value="" disabled>Select FY...</option>
                    {fyOptions.map(fy => (
                        <option key={fy} value={fy}>{fy}</option>
                    ))}
                </select>
            </div>

            {/* Date Inputs */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1.25rem', marginBottom: '1.25rem' }}>
                <div className="form-group">
                    <label className="form-label">From Date</label>
                    <input
                        type="date"
                        className="form-input"
                        value={fromDate}
                        onChange={e => setFromDate(e.target.value)}
                    />
                </div>
                <div className="form-group">
                    <label className="form-label">To Date <span style={{ color: 'var(--text-faint)', fontWeight: 400 }}>(Audit Date)</span></label>
                    <input
                        type="date"
                        className="form-input"
                        value={toDate}
                        onChange={e => setToDate(e.target.value)}
                    />
                </div>
            </div>

            {/* NPA Provision Query */}
            <div className="form-group" style={{ marginBottom: '1.25rem' }}>
                <label className="form-label">Whether NPA provision is available?</label>
                <div style={{ display: 'flex', gap: '1rem', marginTop: 8 }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: '0.875rem' }}>
                        <input type="radio" name="npaAvailable" checked={npaAvailable === 'yes'} onChange={() => { setNpaAvailable('yes'); setNpaSource('same'); }} /> Yes
                    </label>
                    <label style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: '0.875rem' }}>
                        <input type="radio" name="npaAvailable" checked={npaAvailable === 'no'} onChange={() => { setNpaAvailable('no'); setNpaSource(''); }} /> No
                    </label>
                </div>
            </div>

            {npaAvailable === 'yes' && (
                <div className="form-group animate-fade" style={{ marginBottom: '1.5rem', padding: '12px', background: 'var(--p-sky)', borderRadius: 'var(--r-sm)' }}>
                    <label className="form-label" style={{ color: '#0369A1' }}>Where is the NPA provision data?</label>
                    <div style={{ display: 'flex', gap: '1rem', marginTop: 8 }}>
                        <label style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: '0.875rem', color: '#0369A1' }}>
                            <input type="radio" name="npaSource" checked={npaSource === 'same'} onChange={() => setNpaSource('same')} /> In this file
                        </label>
                        <label style={{ display: 'flex', alignItems: 'center', gap: 6, cursor: 'pointer', fontSize: '0.875rem', color: '#0369A1' }}>
                            <input type="radio" name="npaSource" checked={npaSource === 'separate'} onChange={() => setNpaSource('separate')} /> In separate file
                        </label>
                    </div>
                </div>
            )}

            {/* Period Preview */}
            {canProceed && (
                <div className="alert alert-info" style={{ marginBottom: '1.25rem' }}>
                    <Info size={15} style={{ flexShrink: 0, marginTop: 1 }} />
                    <span>Audit period: <strong>{formatDate(fromDate)}</strong> → <strong>{formatDate(toDate)}</strong></span>
                </div>
            )}

            {/* Notes */}
            <div className="alert alert-warning" style={{ marginBottom: '1.5rem' }}>
                <Info size={15} style={{ flexShrink: 0, marginTop: 1 }} />
                <span>The <strong>To Date</strong> is used as the benchmark for NPA days and Sanction Type (New/Old) classification.</span>
            </div>

            <button
                className="btn btn-primary"
                onClick={() => onProceed({ fromDate, toDate, npaAvailable: npaAvailable === 'yes', npaSource })}
                disabled={!canProceed}
                style={{ width: '100%' }}
            >
                Proceed to Upload <ArrowRight size={16} />
            </button>
        </div>
    );
}
