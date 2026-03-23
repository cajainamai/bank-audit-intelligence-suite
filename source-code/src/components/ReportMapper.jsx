import React, { useState, useEffect, useRef } from 'react';
import { ArrowLeft, Download, Upload, Settings2, Zap, Info } from 'lucide-react';

// eslint-disable-next-line react-refresh/only-export-components
export const TARGET_FIELDS = [
    "CIF",
    "A/c No.",
    "Borrower Name",
    "Loan Type",
    "ROI",
    "Sanction Limit",
    "Net O/s",
    "Overdrawn (Yes/No)",
    "Credit Balance",
    "Sanction Date",
    "Sanction Type",
    "Limit Expiry Date",
    "A/c Open Date",
    "PAN",
    "Entity Type",
    "Aadhar No.",
    "Facility Code",
    "PS Flag",
    "Drawing Power",
    "Primary Security",
    "Collateral Security",
    "Asset Class",
    "Asset Classification",
    "NPA Date",
    "Overdue Amount",
    "Outstanding EMIs",
    "NPA Provision"
];

const CALCULATED_FIELDS = ["Loan Type", "Overdrawn (Yes/No)", "Sanction Type", "Entity Type", "Asset Classification", "Outstanding EMIs"];

// Group fields for easier scanning
const FIELD_GROUPS = [
    { label: 'Account Identity', color: 'var(--p-lavender)', badge: 'badge-lavender', fields: ['CIF', 'A/c No.', 'Borrower Name', 'PAN', 'Aadhar No.', 'Entity Type'] },
    { label: 'Facility Details', color: 'var(--p-sky)', badge: 'badge-sky', fields: ['Loan Type', 'Facility Code', 'Sanction Type', 'ROI', 'PS Flag'] },
    { label: 'Amounts & Balances', color: 'var(--p-mint)', badge: 'badge-mint', fields: ['Sanction Limit', 'Net O/s', 'Credit Balance', 'Drawing Power', 'Overdrawn (Yes/No)', 'Overdue Amount', 'Outstanding EMIs', 'NPA Provision'] },
    { label: 'Dates', color: 'var(--p-peach)', badge: 'badge-peach', fields: ['Sanction Date', 'Limit Expiry Date', 'A/c Open Date', 'NPA Date'] },
    { label: 'Security & Classification', color: 'var(--p-lilac)', badge: 'badge-lavender', fields: ['Primary Security', 'Collateral Security', 'Asset Class', 'Asset Classification'] },
];

export default function ReportMapper({ availableColumns, onGenerate, onBack, npaConfig, npaFileDetails }) {
    const [fieldMappings, setFieldMappings] = useState({});
    const [npaMappings, setNpaMappings] = useState({ joinCol: '', provisionCol: '' });
    const [generating, setGenerating] = useState(false);
    const fileInputRef = useRef(null);

    const isSeparateNpa = npaConfig?.available && npaConfig?.source === 'separate';
    const showNpaField = npaConfig?.available;

    useEffect(() => {
        const saved = localStorage.getItem('dataSummarizer_fieldMappings');
        if (saved) {
            try {
                setFieldMappings(JSON.parse(saved));
            } catch (e) { /* ignore */ }
        } else {
            const initialMapping = {};
            const lowerCols = availableColumns.map(c => c.toLowerCase());
            TARGET_FIELDS.forEach(field => {
                if (CALCULATED_FIELDS.includes(field)) {
                    initialMapping[field] = 'AUTO_CALCULATED';
                } else {
                    const lowerField = field.toLowerCase();
                    const matchIndex = lowerCols.findIndex(c => c === lowerField || c.includes(lowerField) || lowerField.includes(c));
                    initialMapping[field] = matchIndex !== -1 ? availableColumns[matchIndex] : '';
                }
            });
            setFieldMappings(initialMapping);
        }

        const savedNpa = localStorage.getItem('dataSummarizer_npaMappings');
        if (savedNpa) {
            try {
                setNpaMappings(JSON.parse(savedNpa));
            } catch (e) { /* ignore */ }
        }
    }, [availableColumns]);

    useEffect(() => {
        if (Object.keys(fieldMappings).length > 0) {
            localStorage.setItem('dataSummarizer_fieldMappings', JSON.stringify(fieldMappings));
        }
    }, [fieldMappings]);

    useEffect(() => {
        if (npaMappings.joinCol || npaMappings.provisionCol) {
            localStorage.setItem('dataSummarizer_npaMappings', JSON.stringify(npaMappings));
        }
    }, [npaMappings]);

    const handleChange = (field, col) => setFieldMappings(prev => ({ ...prev, [field]: col }));

    const handleGenerate = () => {
        setGenerating(true);
        onGenerate(fieldMappings, npaMappings);
    };

    const handleExportConfig = () => {
        const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(fieldMappings, null, 2));
        const dt = new Date();
        const a = document.createElement('a');
        a.href = dataStr;
        a.download = `audit_config_${dt.getFullYear()}${String(dt.getMonth() + 1).padStart(2, '0')}${String(dt.getDate()).padStart(2, '0')}.json`;
        a.click();
    };

    const handleImportConfig = (e) => {
        const fr = new FileReader();
        if (e.target.files[0]) {
            fr.readAsText(e.target.files[0], "UTF-8");
            fr.onload = (event) => {
                try { setFieldMappings(JSON.parse(event.target.result)); }
                catch { alert("Invalid config file."); }
            };
            e.target.value = null;
        }
    };

    const mappedCount = TARGET_FIELDS.filter(f => !CALCULATED_FIELDS.includes(f) && fieldMappings[f]).length;
    const totalMappable = TARGET_FIELDS.filter(f => !CALCULATED_FIELDS.includes(f)).length;

    return (
        <div className="card card-body animate-slide" style={{ maxWidth: 800, margin: '0 auto' }}>
            {/* Header */}
            <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: '1rem', marginBottom: '1.5rem' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                    <div style={{
                        width: 44, height: 44,
                        background: 'var(--p-lavender)',
                        borderRadius: 11,
                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                        flexShrink: 0
                    }}>
                        <Settings2 size={20} color="#4C1D95" />
                    </div>
                    <div>
                        <h2 className="page-title" style={{ fontSize: '1.375rem' }}>Map Report Fields</h2>
                        <p style={{ fontSize: '0.8375rem', color: 'var(--text-muted)', marginTop: 2 }}>
                            {mappedCount} / {totalMappable} fields mapped
                        </p>
                    </div>
                </div>
                <div style={{ display: 'flex', gap: '0.5rem' }}>
                    <button className="btn btn-secondary btn-sm" onClick={() => fileInputRef.current?.click()}>
                        <Upload size={13} /> Import
                    </button>
                    <input type="file" accept=".json" ref={fileInputRef} style={{ display: 'none' }} onChange={handleImportConfig} />
                    <button className="btn btn-secondary btn-sm" onClick={handleExportConfig}>
                        <Download size={13} /> Export
                    </button>
                </div>
            </div>

            <div className="alert alert-info" style={{ marginBottom: '1.5rem' }}>
                <Info size={14} style={{ flexShrink: 0, marginTop: 1 }} />
                <span>Fields marked <strong>Auto-Calculated</strong> are derived automatically. Mappings are auto-saved to your browser for next time.</span>
            </div>

            {/* Field groups */}
            <div style={{ maxHeight: '55vh', overflowY: 'auto', paddingRight: '0.5rem', marginBottom: '1.5rem' }}>
                {FIELD_GROUPS.map(({ label, color, badge, fields }) => (
                    <div key={label} style={{ marginBottom: '1.25rem' }}>
                        <div style={{
                            fontSize: '0.78rem',
                            fontWeight: 700,
                            color: 'var(--text-muted)',
                            textTransform: 'uppercase',
                            letterSpacing: '0.06em',
                            padding: '6px 10px',
                            background: color,
                            borderRadius: 'var(--r-sm)',
                            marginBottom: '0.5rem',
                            display: 'inline-block'
                        }}>
                            {label}
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column' }}>
                            {fields.map(field => {
                                const isCalc = CALCULATED_FIELDS.includes(field);
                                if (field === 'NPA Provision' && !showNpaField) return null;
                                return (
                                    <div key={field} className="mapping-row">
                                        <div className="mapping-field-name">
                                            {field}
                                            {isCalc && <span className="mapping-calc-badge"><Zap size={10} /> Auto</span>}
                                        </div>
                                        <div>
                                            {isCalc ? (
                                                <div style={{
                                                    padding: '8px 12px',
                                                    background: 'var(--p-mint)',
                                                    border: '1px solid var(--p-mint-d)',
                                                    borderRadius: 'var(--r-sm)',
                                                    color: '#064E3B',
                                                    fontSize: '0.8125rem',
                                                    fontWeight: 500
                                                }}>
                                                    System Calculated
                                                </div>
                                            ) : (
                                                <select
                                                    className="form-input"
                                                    value={fieldMappings[field] || ''}
                                                    onChange={e => handleChange(field, e.target.value)}
                                                    style={{
                                                        fontSize: '0.875rem',
                                                        padding: '8px 36px 8px 12px',
                                                        background: field === 'NPA Provision' && isSeparateNpa && npaMappings.joinCol && npaMappings.provisionCol ? 'var(--p-mint)' : 'white',
                                                        fontWeight: field === 'NPA Provision' && isSeparateNpa && npaMappings.joinCol && npaMappings.provisionCol ? 600 : 400
                                                    }}
                                                    disabled={field === 'NPA Provision' && isSeparateNpa}
                                                >
                                                    <option value="">
                                                        {field === 'NPA Provision' && isSeparateNpa
                                                            ? (npaMappings.joinCol && npaMappings.provisionCol ? '✓ Mapped from Separate File' : '— Pending Separate Mapping —')
                                                            : '— Not mapped —'}
                                                    </option>
                                                    {availableColumns.map(col => (
                                                        <option key={col} value={col}>{col}</option>
                                                    ))}
                                                </select>
                                            )}
                                        </div>
                                    </div>
                                );
                            })}
                        </div>
                    </div>
                ))}

                {isSeparateNpa && (
                    <div style={{ marginTop: '2.5rem', borderTop: '2px dashed var(--p-lavender)', paddingTop: '1.5rem' }}>
                        <div style={{
                            fontSize: '0.78rem',
                            fontWeight: 700,
                            color: 'var(--text-muted)',
                            textTransform: 'uppercase',
                            letterSpacing: '0.06em',
                            padding: '6px 10px',
                            background: 'var(--p-sky)',
                            borderRadius: 'var(--r-sm)',
                            marginBottom: '1rem',
                            display: 'inline-block'
                        }}>
                            NPA Provision Mapping (Separate File)
                        </div>
                        <div className="alert alert-info" style={{ marginBottom: '1.25rem' }}>
                            <Info size={14} style={{ flexShrink: 0 }} />
                            <span>Select columns from <strong>{npaFileDetails?.file?.name || 'NPA File'}</strong> to map provisions.</span>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                            <div className="mapping-row">
                                <div className="mapping-field-name">Account ID / Number <span style={{ color: 'red' }}>*</span></div>
                                <select className="form-input" value={npaMappings.joinCol} onChange={e => setNpaMappings(p => ({ ...p, joinCol: e.target.value }))}>
                                    <option value="">— Select Join Column —</option>
                                    {npaFileDetails.headers.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                            <div className="mapping-row">
                                <div className="mapping-field-name">Provision Amount <span style={{ color: 'red' }}>*</span></div>
                                <select className="form-input" value={npaMappings.provisionCol} onChange={e => setNpaMappings(p => ({ ...p, provisionCol: e.target.value }))}>
                                    <option value="">— Select Provision Column —</option>
                                    {npaFileDetails.headers.map(h => <option key={h} value={h}>{h}</option>)}
                                </select>
                            </div>
                        </div>
                    </div>
                )}
            </div>

            {/* Actions */}
            <hr className="divider" />
            <div style={{ display: 'flex', gap: '0.75rem', justifyContent: 'space-between', alignItems: 'center' }}>
                <button className="btn btn-secondary" onClick={onBack}><ArrowLeft size={15} /> Back</button>
                <button
                    className="btn btn-success btn-lg"
                    onClick={handleGenerate}
                    disabled={generating}
                    style={{ minWidth: 200 }}
                >
                    <Download size={17} /> Download Excel Report
                </button>
            </div>
        </div>
    );
}
