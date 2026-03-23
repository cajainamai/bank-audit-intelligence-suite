import React, { useState } from 'react';
import { Download, ArrowLeft, ChevronLeft, ChevronRight } from 'lucide-react';

export default function ResultsGrid({ data, columns, onDownload, onBack }) {
    const [page, setPage] = useState(0);
    const rowsPerPage = 50;

    const totalPages = Math.ceil(data.length / rowsPerPage);
    const currentData = data.slice(page * rowsPerPage, (page + 1) * rowsPerPage);

    return (
        <div className="card glass-panel" style={{ maxWidth: '100%', margin: '0 auto', display: 'flex', flexDirection: 'column', height: '80vh' }}>
            <div style={{ padding: '20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid rgba(0,0,0,0.1)' }}>
                <div>
                    <h2 className="page-title" style={{ fontSize: '1.8rem', margin: 0 }}>Audit Report Preview</h2>
                    <p className="page-subtitle" style={{ margin: 0, marginTop: '5px' }}>
                        Showing {data.length} records. Review your data before downloading.
                    </p>
                </div>
                <div style={{ display: 'flex', gap: '1rem' }}>
                    <button className="btn btn-secondary" onClick={onBack}>
                        <ArrowLeft size={18} /> Back
                    </button>
                    <button className="btn btn-primary" onClick={onDownload}>
                        <Download size={18} /> Download Excel
                    </button>
                </div>
            </div>

            <div style={{ flex: 1, overflow: 'auto' }}>
                <table style={{ minWidth: '100%', borderCollapse: 'collapse', fontSize: '0.9rem' }}>
                    <thead style={{ position: 'sticky', top: 0, zIndex: 1, backgroundColor: 'var(--pastel-lavender)' }}>
                        <tr>
                            {columns.map((col, idx) => (
                                <th key={idx} style={{ padding: '12px 16px', textAlign: 'left', borderBottom: '2px solid rgba(0,0,0,0.1)', color: 'var(--text-dark)', whiteSpace: 'nowrap' }}>
                                    {col}
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {currentData.map((row, rIdx) => (
                            <tr key={rIdx} style={{ backgroundColor: rIdx % 2 === 0 ? 'white' : 'var(--background-color)' }}>
                                {columns.map((col, cIdx) => {
                                    let val = row[col];
                                    // Formatting for preview UI specifically
                                    if (val instanceof Date) {
                                        val = val.toLocaleDateString('en-GB'); // DD/MM/YYYY
                                    } else if (typeof val === 'number') {
                                        if (col === 'ROI') {
                                            val = (val * 100).toFixed(2) + '%';
                                        } else {
                                            val = val.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                                        }
                                    }
                                    return (
                                        <td key={cIdx} style={{ padding: '10px 16px', borderBottom: '1px solid rgba(0,0,0,0.05)', whiteSpace: 'nowrap' }}>
                                            {val}
                                        </td>
                                    );
                                })}
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>

            {totalPages > 1 && (
                <div style={{ padding: '15px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: '1px solid rgba(0,0,0,0.1)' }}>
                    <span style={{ fontSize: '0.9rem', color: '#666' }}>
                        Showing {page * rowsPerPage + 1} to {Math.min((page + 1) * rowsPerPage, data.length)} of {data.length} entries
                    </span>
                    <div style={{ display: 'flex', gap: '10px' }}>
                        <button
                            className="btn btn-secondary"
                            style={{ padding: '6px 12px' }}
                            onClick={() => setPage(p => Math.max(0, p - 1))}
                            disabled={page === 0}
                        >
                            <ChevronLeft size={16} /> Previous
                        </button>
                        <button
                            className="btn btn-secondary"
                            style={{ padding: '6px 12px' }}
                            onClick={() => setPage(p => Math.min(totalPages - 1, p + 1))}
                            disabled={page === totalPages - 1}
                        >
                            Next <ChevronRight size={16} />
                        </button>
                    </div>
                </div>
            )}
        </div>
    );
}
