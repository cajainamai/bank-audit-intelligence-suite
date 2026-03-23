import React, { useCallback, useState } from 'react';
import { UploadCloud, FileText, CheckCircle, X, ArrowRight, Info } from 'lucide-react';

function DropZone({ zoneKey, file, onClear, onDrop, onChange, title, subtitle, accent, required }) {
    const [dragging, setDragging] = useState(false);
    const hasFile = !!file;

    const handleDrag = (e) => {
        e.preventDefault(); e.stopPropagation();
        if (e.type === 'dragenter' || e.type === 'dragover') setDragging(true);
        else setDragging(false);
    };
    const handleDrop = (e) => {
        e.preventDefault(); e.stopPropagation();
        setDragging(false);
        onDrop(e, zoneKey);
    };
    const handleChange = (e) => {
        e.preventDefault();
        onChange(e, zoneKey);
    };

    return (
        <div>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '0.625rem' }}>
                <div>
                    <div style={{ fontWeight: 700, fontSize: '0.9375rem', color: 'var(--text-heading)', display: 'flex', alignItems: 'center', gap: 6 }}>
                        {title}
                        {!required && <span className="badge badge-peach" style={{ fontSize: '0.68rem', padding: '1px 8px' }}>Optional</span>}
                    </div>
                    <div style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginTop: 2 }}>{subtitle}</div>
                </div>
                {hasFile && (
                    <button onClick={onClear} style={{ background: 'none', border: 'none', color: 'var(--text-faint)', cursor: 'pointer', padding: 4, borderRadius: 4 }}>
                        <X size={14} />
                    </button>
                )}
            </div>

            <label
                className={`upload-zone ${dragging ? 'drag-over' : ''} ${hasFile ? 'has-file' : ''}`}
                onDragEnter={handleDrag} onDragLeave={handleDrag} onDragOver={handleDrag} onDrop={handleDrop}
                style={{ cursor: 'pointer', display: 'block', position: 'relative' }}
            >
                {hasFile ? (
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8 }}>
                        <CheckCircle size={32} color="var(--brand-success)" />
                        <span style={{ fontWeight: 600, fontSize: '0.875rem', color: 'var(--text-heading)', textAlign: 'center', wordBreak: 'break-all' }}>
                            {file.name}
                        </span>
                        <span style={{ fontSize: '0.76rem', color: 'var(--text-muted)' }}>Click or drop to replace</span>
                    </div>
                ) : (
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8 }}>
                        {zoneKey === 'data'
                            ? <UploadCloud size={34} color="var(--text-faint)" />
                            : <FileText size={34} color="var(--text-faint)" />
                        }
                        <span style={{ fontWeight: 600, fontSize: '0.875rem', color: 'var(--text-muted)' }}>
                            Drop file here or <span style={{ color: 'var(--brand-primary)' }}>browse</span>
                        </span>
                        <span style={{ fontSize: '0.75rem', color: 'var(--text-faint)' }}>.xlsx &bull; .xls</span>
                    </div>
                )}
                <input type="file" accept=".xlsx,.xls,.csv" onChange={handleChange}
                    style={{ position: 'absolute', inset: 0, opacity: 0, cursor: 'pointer' }} />
            </label>
        </div>
    );
}

export default function FileUpload({ onFileUpload, npaConfig }) {
    const [dataFile, setDataFile] = useState(null);
    const [mappingFile, setMappingFile] = useState(null);
    const [npaFile, setNpaFile] = useState(null);

    const validate = file => file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.name.endsWith('.csv'));

    const handleDrop = useCallback((e, zone) => {
        const file = e.dataTransfer.files?.[0];
        if (validate(file)) {
            if (zone === 'data') setDataFile(file);
            else if (zone === 'mapping') setMappingFile(file);
            else if (zone === 'npa') setNpaFile(file);
        }
    }, []);

    const handleChange = (e, zone) => {
        const file = e.target.files?.[0];
        if (validate(file)) {
            if (zone === 'data') setDataFile(file);
            else if (zone === 'mapping') setMappingFile(file);
            else if (zone === 'npa') setNpaFile(file);
        }
        e.target.value = '';
    };

    const isSeparateNpa = npaConfig?.available && npaConfig?.source === 'separate';

    return (
        <div className="card card-body animate-slide" style={{ maxWidth: 800, margin: '0 auto' }}>
            <div style={{ marginBottom: '1.75rem' }}>
                <h2 className="page-title" style={{ fontSize: '1.375rem' }}>Upload Files</h2>
                <p className="page-subtitle" style={{ marginTop: 4 }}>
                    Upload your Bank CBS data extract and any other relevant mapping/NPA files.
                </p>
            </div>

            <hr className="divider" />

            <div style={{
                display: 'grid',
                gridTemplateColumns: isSeparateNpa ? '1fr 1fr 1fr' : '1fr 1fr',
                gap: '1.25rem',
                marginBottom: '1.5rem'
            }}>
                <DropZone
                    zoneKey="data" file={dataFile}
                    onClear={() => setDataFile(null)}
                    onDrop={handleDrop} onChange={handleChange}
                    title="Bank Data Extract" subtitle="CBS ledger/account export"
                    required
                />
                <DropZone
                    zoneKey="mapping" file={mappingFile}
                    onClear={() => setMappingFile(null)}
                    onDrop={handleDrop} onChange={handleChange}
                    title="Product Mapping" subtitle="Facility code → Loan type"
                    required={false}
                />
                {isSeparateNpa && (
                    <DropZone
                        zoneKey="npa" file={npaFile}
                        onClear={() => setNpaFile(null)}
                        onDrop={handleDrop} onChange={handleChange}
                        title="NPA Provision File" subtitle="Account wise provision data"
                        accent="var(--p-sky)"
                        required
                    />
                )}
            </div>

            <div className="alert alert-info" style={{ marginBottom: '1.5rem' }}>
                <Info size={15} style={{ flexShrink: 0, marginTop: 1 }} />
                <span>
                    {isSeparateNpa
                        ? "Please upload the 3rd Excel file containing account-wise NPA provisions."
                        : "The Product Mapping file maps facility/product codes to loan types (e.g., TL, CC, OD, Gold)."}
                </span>
            </div>

            <button
                className="btn btn-primary"
                onClick={() => dataFile && onFileUpload(dataFile, mappingFile, npaFile)}
                disabled={!dataFile || (isSeparateNpa && !npaFile)}
                style={{ width: '100%' }}
            >
                Analyse Data <ArrowRight size={16} />
            </button>
        </div>
    );
}
